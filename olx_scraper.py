from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Callable, Dict, Iterable, List, Optional

import requests
from bs4 import BeautifulSoup
from slugify import slugify

USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/122.0 Safari/537.36"
)

MONTHS = {
    "янв": 1,
    "января": 1,
    "фев": 2,
    "февр": 2,
    "февраля": 2,
    "мар": 3,
    "марта": 3,
    "апр": 4,
    "апреля": 4,
    "май": 5,
    "мая": 5,
    "июн": 6,
    "июня": 6,
    "июл": 7,
    "июля": 7,
    "авг": 8,
    "августа": 8,
    "сен": 9,
    "сент": 9,
    "сентября": 9,
    "окт": 10,
    "октября": 10,
    "ноя": 11,
    "ноября": 11,
    "дек": 12,
    "декабря": 12,
    # Ukrainian
    "січ": 1,
    "січня": 1,
    "лют": 2,
    "лютого": 2,
    "бер": 3,
    "березня": 3,
    "кві": 4,
    "квітня": 4,
    "тра": 5,
    "травня": 5,
    "чер": 6,
    "червня": 6,
    "лип": 7,
    "липня": 7,
    "сер": 8,
    "серпня": 8,
    "вер": 9,
    "вересня": 9,
    "жов": 10,
    "жовтня": 10,
    "лис": 11,
    "листопада": 11,
    "гру": 12,
    "грудня": 12,
}


@dataclass
class Listing:
    query: str
    title: str
    price: Optional[int]
    currency: Optional[str]
    location: Optional[str]
    posted_at: Optional[datetime]
    url: str
    image_url: Optional[str]
    listing_id: str
    description: Optional[str] = None
    sheet_name: Optional[str] = None
    seller_rating: Optional[float] = None
    seller_reviews: Optional[int] = None
    seller_score: Optional[float] = None
    matched_tags: Optional[str] = None


def _parse_date(text: str, now: datetime) -> Optional[datetime]:
    if not text:
        return None
    raw = text.strip().lower()

    if "сегодня" in raw or "сьогодні" in raw:
        return now
    if "вчера" in raw or "вчора" in raw:
        return now - timedelta(days=1)

    # Example: "18 февр." or "18 лютого"
    match = re.search(r"(\d{1,2})\s+([а-яіїєё\.]+)", raw)
    if not match:
        return None

    day = int(match.group(1))
    month_raw = match.group(2).replace(".", "")
    month = MONTHS.get(month_raw)
    if not month:
        return None

    year = now.year
    try_date = datetime(year, month, day)
    if try_date > now + timedelta(days=1):
        try_date = datetime(year - 1, month, day)
    return try_date


def _extract_price(text: str) -> (Optional[int], Optional[str]):
    if not text:
        return None, None
    currency = None
    if "грн" in text.lower():
        currency = "UAH"
    elif "$" in text:
        currency = "USD"

    digits = re.sub(r"[^0-9]", "", text)
    if not digits:
        return None, currency
    return int(digits), currency


def _listing_id_from_url(url: str) -> str:
    match = re.search(r"-ID(\d+)", url)
    if match:
        return match.group(1)
    match = re.search(r"/(\d+)/?$", url)
    if match:
        return match.group(1)
    return url


def _build_search_url(query: str, price_from: Optional[int], price_to: Optional[int], city: Optional[str]) -> str:
    query_slug = slugify(query, lowercase=True)

    if city:
        city_slug = slugify(city, lowercase=True)
        base = f"https://www.olx.ua/d/uk/{city_slug}/q-{query_slug}/"
    else:
        base = f"https://www.olx.ua/d/uk/list/q-{query_slug}/"

    params = []
    if price_from is not None:
        params.append(f"search%5Bfilter_float_price%3Afrom%5D={price_from}")
    if price_to is not None:
        params.append(f"search%5Bfilter_float_price%3Ato%5D={price_to}")
    if params:
        return base + "?" + "&".join(params)
    return base


def _get_cards(soup: BeautifulSoup) -> List:
    cards = soup.select('[data-cy="l-card"]')
    if cards:
        return cards
    cards = soup.select('a[data-cy="ad-card"]')
    if cards:
        return cards
    cards = soup.select('a[href*="/d/uk/"]')
    return cards


def _extract_listing(card, query: str, now: datetime) -> Optional[Listing]:
    link = None
    if getattr(card, "name", "") == "a":
        link = card.get("href")
    else:
        link_tag = card.find("a", href=True)
        if link_tag:
            link = link_tag.get("href")
    if not link:
        return None
    if link.startswith("/"):
        link = "https://www.olx.ua" + link

    title = None
    for sel in [
        "h6",
        "h4",
        "[data-cy='ad-card-title']",
        "[data-testid='ad-title']",
    ]:
        t = card.select_one(sel)
        if t and t.get_text(strip=True):
            title = t.get_text(strip=True)
            break
    if not title:
        title = card.get_text(" ", strip=True)[:120]

    price_text = None
    for sel in [
        "[data-testid='ad-price']",
        "[data-cy='ad-card-price']",
        "p",
        "span",
    ]:
        t = card.select_one(sel)
        if t and "грн" in t.get_text().lower():
            price_text = t.get_text(strip=True)
            break

    price, currency = _extract_price(price_text or "")

    location = None
    posted_at = None
    loc_tag = card.select_one("[data-testid='location-date']")
    if loc_tag:
        loc_text = loc_tag.get_text(" ", strip=True)
        parts = [p.strip() for p in loc_text.split(",") if p.strip()]
        if parts:
            if len(parts) >= 2:
                location = parts[0]
                posted_at = _parse_date(parts[-1], now)
            else:
                posted_at = _parse_date(parts[0], now)

    image_url = None
    img = card.find("img")
    if img:
        image_url = img.get("src") or img.get("data-src")

    listing_id = _listing_id_from_url(link)

    return Listing(
        query=query,
        title=title,
        price=price,
        currency=currency,
        location=location,
        posted_at=posted_at,
        url=link,
        image_url=image_url,
        listing_id=listing_id,
    )


def _extract_aggregate_rating(soup: BeautifulSoup) -> (Optional[float], Optional[int]):
    scripts = soup.find_all("script", attrs={"type": "application/ld+json"})
    for s in scripts:
        try:
            data = json.loads(s.string or "")
        except Exception:
            continue
        items = data if isinstance(data, list) else [data]
        for item in items:
            if not isinstance(item, dict):
                continue
            ar = item.get("aggregateRating")
            if isinstance(ar, dict):
                rating = ar.get("ratingValue")
                reviews = ar.get("reviewCount")
                try:
                    rating_f = float(rating) if rating is not None else None
                except Exception:
                    rating_f = None
                try:
                    reviews_i = int(reviews) if reviews is not None else None
                except Exception:
                    reviews_i = None
                if rating_f is not None or reviews_i is not None:
                    return rating_f, reviews_i
    return None, None


def _fetch_details(url: str, headers: Dict[str, str]) -> (Optional[str], Optional[float], Optional[int]):
    try:
        resp = requests.get(url, headers=headers, timeout=30)
    except requests.RequestException:
        return None, None, None
    if resp.status_code != 200:
        return None, None, None
    soup = BeautifulSoup(resp.text, "html.parser")
    rating, reviews = _extract_aggregate_rating(soup)
    for sel in [
        "[data-cy='ad_description']",
        "[data-testid='ad-description']",
        "div[data-testid='ad-description']",
        "div[data-cy='ad_description']",
    ]:
        node = soup.select_one(sel)
        if node:
            text = node.get_text("\n", strip=True)
            if text:
                return text, rating, reviews
    return None, rating, reviews


def search_olx(
    query: str,
    price_from: Optional[int],
    price_to: Optional[int],
    city: Optional[str],
    pages: int = 3,
    max_age_days: int = 92,
    photos_only: bool = True,
    fetch_details: bool = True,
    progress_cb: Optional[Callable[[str], None]] = None,
    page_delay: float = 0.0,
    detail_delay: float = 0.0,
    max_listings: Optional[int] = None,
) -> List[Listing]:
    now = datetime.now()
    cutoff = now - timedelta(days=max_age_days)
    headers = {"User-Agent": USER_AGENT}

    results: List[Listing] = []
    reported = 0
    for page in range(1, pages + 1):
        if progress_cb:
            progress_cb(f"Page {page}/{pages} started")
        url = _build_search_url(query, price_from, price_to, city)
        if page > 1:
            join_char = "&" if "?" in url else "?"
            url = f"{url}{join_char}page={page}"

        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code != 200:
            break
        soup = BeautifulSoup(resp.text, "html.parser")
        cards = _get_cards(soup)
        if not cards:
            break

        for card in cards:
            if max_listings is not None and len(results) >= max_listings:
                return results
            listing = _extract_listing(card, query, now)
            if not listing:
                continue
            if photos_only and not listing.image_url:
                continue
            if listing.posted_at and listing.posted_at < cutoff:
                continue
            if fetch_details:
                description, rating, reviews = _fetch_details(listing.url, headers)
                listing.description = description
                listing.seller_rating = rating
                listing.seller_reviews = reviews
                if rating is not None and reviews is not None:
                    listing.seller_score = rating * reviews
                if detail_delay > 0:
                    time.sleep(detail_delay)
            results.append(listing)
            if progress_cb:
                if max_listings is not None:
                    if len(results) == max_listings or len(results) - reported >= 10:
                        reported = len(results)
                        progress_cb(f"Page {page}/{pages}: processed {len(results)}/{max_listings} listings")
                else:
                    if len(results) - reported >= 10:
                        reported = len(results)
                        progress_cb(f"Page {page}/{pages}: processed {len(results)} listings")

        if page_delay > 0:
            time.sleep(page_delay)

    return results
