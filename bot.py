from __future__ import annotations

import asyncio
import hashlib
import logging
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Set

from deep_translator import GoogleTranslator

import pandas as pd
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from dotenv import load_dotenv
from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters
from zoneinfo import ZoneInfo

from olx_scraper import Listing, search_olx
from storage import (
    ensure_dir,
    load_allowed_users,
    load_queries,
    load_seen_ids,
    save_allowed_users,
    save_queries,
    save_seen_ids,
)

load_dotenv()

DATA_DIR = Path(os.getenv("DATA_DIR", "data"))
TIMEZONE = os.getenv("TIMEZONE", "Europe/Kyiv")
CHAT_ID_ENV = os.getenv("TELEGRAM_CHAT_ID")
SEARCH_PAGES = int(os.getenv("SEARCH_PAGES", "3"))
PAGE_DELAY = float(os.getenv("PAGE_DELAY", "0"))
DETAIL_DELAY = float(os.getenv("DETAIL_DELAY", "0"))
MAX_LISTINGS_PER_QUERY = os.getenv("MAX_LISTINGS_PER_QUERY")
MAX_LISTINGS_PER_QUERY = int(MAX_LISTINGS_PER_QUERY) if MAX_LISTINGS_PER_QUERY else None
QUERY_COOLDOWN_MINUTES = int(os.getenv("QUERY_COOLDOWN_MINUTES", "30"))
LOG_FILE = Path(os.getenv("LOG_FILE", "data/bot.log"))
ALLOWED_USERS_FILE = Path(os.getenv("ALLOWED_USERS_FILE", "data/allowed_users.json"))
ADMIN_CHAT_IDS = {
    int(x.strip()) for x in os.getenv("ADMIN_CHAT_IDS", "").split(",") if x.strip()
}

logger = logging.getLogger("olx-bot")

HELP_TEXT = (
    "Отправьте запрос в формате:\n"
    "Название, от 1000 до 50000, Город\n"
    "или без города:\n"
    "Название, от 1000 до 50000\n\n"
    "Команды:\n"
    "/help — список команд\n"
    "/list — список запросов\n"
    "/remove N — удалить запрос по номеру\n"
    "/run — выполнить сбор сейчас\n"
    "/report — отправить Excel сейчас\n"
    "/log — последние строки лога\n"
    "/myid — показать ваш Telegram user_id\n"
    "/search <запрос> — добавить запрос (пример: /search Дом, от 100000 до 4000000, г. Одесса)"
)


def _get_user_dir(user_id: int) -> Path:
    return DATA_DIR / "users" / str(user_id)


def _query_file(user_id: int) -> Path:
    return _get_user_dir(user_id) / "queries.json"


def _seen_file(user_id: int) -> Path:
    return _get_user_dir(user_id) / "seen_ids.json"


def _excel_file(user_id: int) -> Path:
    return _get_user_dir(user_id) / "olx.xlsx"


def _chat_id_file(user_id: int) -> Path:
    return _get_user_dir(user_id) / "chat_id.txt"


def _load_chat_id(user_id: int) -> Optional[int]:
    if CHAT_ID_ENV:
        return int(CHAT_ID_ENV)
    chat_file = _chat_id_file(user_id)
    if chat_file.exists():
        return int(chat_file.read_text(encoding="utf-8").strip())
    return None


def _save_chat_id(user_id: int, chat_id: int) -> None:
    user_dir = _get_user_dir(user_id)
    ensure_dir(user_dir)
    _chat_id_file(user_id).write_text(str(chat_id), encoding="utf-8")


def _is_admin(user_id: int) -> bool:
    return user_id in ADMIN_CHAT_IDS


def _allowed_users() -> Set[int]:
    ensure_dir(DATA_DIR)
    return load_allowed_users(ALLOWED_USERS_FILE)


def _is_allowed(user_id: int) -> bool:
    if _is_admin(user_id):
        return True
    return user_id in _allowed_users()


def _parse_query(text: str) -> Optional[Dict]:
    raw = text.strip()
    if not raw:
        return None

    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if not parts:
        return None

    query = parts[0]
    price_from = None
    price_to = None
    city = None

    price_text = ",".join(parts[1:]) if len(parts) > 1 else ""

    match = re.search(r"от\s*(\d+)\s*до\s*(\d+)", price_text, re.IGNORECASE)
    if match:
        price_from = int(match.group(1))
        price_to = int(match.group(2))
        if len(parts) >= 3:
            city = parts[2]
    else:
        if len(parts) >= 2:
            city = parts[1]

    return {
        "query": query,
        "price_from": price_from,
        "price_to": price_to,
        "city": city,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }


def _translate_ru_uk(text: str) -> Optional[str]:
    try:
        return GoogleTranslator(source="auto", target="uk").translate(text)
    except Exception:
        return None


def _append_to_excel(user_id: int, listings: List[Listing]) -> int:
    if not listings:
        return 0

    user_dir = _get_user_dir(user_id)
    ensure_dir(user_dir)
    excel_path = _excel_file(user_id)

    rows_by_sheet = {}
    for l in listings:
        sheet_name = l.sheet_name or "Sheet1"
        rows_by_sheet.setdefault(sheet_name, []).append(
            {
                "scraped_at": datetime.now().isoformat(timespec="seconds"),
                "query": l.query,
                "title": l.title,
                "price": l.price,
                "currency": l.currency,
                "location": l.location,
                "posted_at": l.posted_at.isoformat(timespec="seconds") if l.posted_at else None,
                "url": l.url,
                "image_url": l.image_url,
                "listing_id": l.listing_id,
                "description": l.description,
                "seller_rating": l.seller_rating,
                "seller_reviews": l.seller_reviews,
                "seller_score": l.seller_score,
            }
        )

    if excel_path.exists():
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            for sheet_name, rows in rows_by_sheet.items():
                df = pd.DataFrame(rows)
                sheet = writer.sheets.get(sheet_name)
                start_row = sheet.max_row if sheet else 0
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=start_row == 0, startrow=start_row)
    else:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            for sheet_name, rows in rows_by_sheet.items():
                df = pd.DataFrame(rows)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    _format_excel(excel_path)
    return sum(len(r) for r in rows_by_sheet.values())


def _prune_excel(excel_path: Path, max_age_days: int = 40) -> int:
    if not excel_path.exists():
        return 0
    try:
        xls = pd.ExcelFile(excel_path)
    except Exception:
        return 0
    if not xls.sheet_names:
        return 0
    now = datetime.now()
    cutoff = now - pd.Timedelta(days=max_age_days)
    total_removed = 0
    sheets_out = {}
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name)
        if df.empty:
            sheets_out[name] = df
            continue
        posted = pd.to_datetime(df.get("posted_at"), errors="coerce")
        scraped = pd.to_datetime(df.get("scraped_at"), errors="coerce")
        effective = posted.fillna(scraped)
        keep = effective.isna() | (effective >= cutoff)
        pruned = df[keep].copy()
        total_removed += len(df) - len(pruned)
        sheets_out[name] = pruned
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for name, df in sheets_out.items():
            df.to_excel(writer, sheet_name=name, index=False)
    _format_excel(excel_path)
    return total_removed


def _load_recent_from_excel(excel_path: Path, hours: int = 24) -> List[Listing]:
    if not excel_path.exists():
        return []
    try:
        xls = pd.ExcelFile(excel_path)
    except Exception:
        return []
    frames = []
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name)
        if df.empty:
            continue
        df["__sheet"] = name
        frames.append(df)
    if not frames:
        return []
    df = pd.concat(frames, ignore_index=True)
    if "scraped_at" not in df.columns:
        return []
    df["scraped_at"] = pd.to_datetime(df["scraped_at"], errors="coerce")
    cutoff = datetime.now() - pd.Timedelta(hours=hours)
    recent = df[df["scraped_at"] >= cutoff].copy()
    listings = []
    for _, row in recent.iterrows():
        listings.append(
            Listing(
                query=row.get("query", ""),
                title=row.get("title", ""),
                price=int(row["price"]) if pd.notna(row.get("price")) else None,
                currency=row.get("currency", None),
                location=row.get("location", None),
                posted_at=pd.to_datetime(row.get("posted_at"), errors="coerce").to_pydatetime()
                if pd.notna(row.get("posted_at"))
                else None,
                url=row.get("url", ""),
                image_url=row.get("image_url", None),
                listing_id=str(row.get("listing_id", "")),
                description=row.get("description", None),
                seller_rating=row.get("seller_rating", None),
                seller_reviews=row.get("seller_reviews", None),
                seller_score=row.get("seller_score", None),
                sheet_name=row.get("__sheet", None),
            )
        )
    return listings


def _sheet_name_for_query(q: Dict) -> str:
    base = q.get("query", "").strip()
    cleaned = re.sub(r"[\\/:*?\[\]]", " ", base).strip()
    if len(cleaned) <= 31:
        return cleaned or "Sheet"
    digest = hashlib.md5(base.encode("utf-8")).hexdigest()[:6]
    trimmed = cleaned[:24].rstrip()
    return f"{trimmed}~{digest}"


def _format_excel(excel_path: Path) -> None:
    if not excel_path.exists():
        return
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font
    except Exception:
        return
    wb = load_workbook(excel_path)
    header_font = Font(bold=True)
    wrap = Alignment(wrap_text=True, vertical="top")
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = header_font
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = wrap
        for col_cells in ws.columns:
            max_len = 0
            col = col_cells[0].column_letter
            for cell in col_cells[:500]:
                if cell.value is None:
                    continue
                max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col].width = min(max(10, max_len + 2), 60)
    wb.save(excel_path)


def _format_summary(listings: List[Listing], limit: int = 5) -> str:
    if not listings:
        return "Новых объявлений: 0"
    lines = [f"Новых объявлений: {len(listings)}", "Топ:"]
    for l in listings[:limit]:
        price = f"{l.price} {l.currency}" if l.price else "Цена не указана"
        location = l.location or "Локация не указана"
        lines.append(f"- {l.title} | {price} | {location}")
    return "\n".join(lines)


def _user_from_update(update: Update) -> Optional[int]:
    if update.effective_user:
        return update.effective_user.id
    return None


def _user_allowed(update: Update) -> bool:
    user_id = _user_from_update(update)
    if user_id is None:
        return False
    return _is_allowed(user_id)


def _ensure_user_dir(user_id: int) -> None:
    ensure_dir(_get_user_dir(user_id))


def _allowed_user_list() -> Set[int]:
    return _allowed_users()


def _save_allowed_user_list(users: Set[int]) -> None:
    save_allowed_users(ALLOWED_USERS_FILE, users)


def _add_allowed_user(user_id: int) -> None:
    users = _allowed_user_list()
    users.add(user_id)
    _save_allowed_user_list(users)


def _remove_allowed_user(user_id: int) -> None:
    users = _allowed_user_list()
    if user_id in users:
        users.remove(user_id)
    _save_allowed_user_list(users)


def _current_users() -> List[int]:
    return sorted(_allowed_user_list())


def _is_admin_update(update: Update) -> bool:
    user_id = _user_from_update(update)
    if user_id is None:
        return False
    return _is_admin(user_id)


async def _run_scrape_for_user(user_id: int, app: Application, force: bool = False) -> int:
    user_dir = _get_user_dir(user_id)
    ensure_dir(user_dir)

    queries = load_queries(_query_file(user_id))
    if not queries:
        logger.info("No queries for user %s; skipping scrape.", user_id)
        return 0

    seen = load_seen_ids(_seen_file(user_id))
    new_listings: List[Listing] = []
    now = datetime.now()

    for q in queries:
        if not force:
            last_run = q.get("last_run")
            if last_run:
                try:
                    last_dt = datetime.fromisoformat(last_run)
                    if (now - last_dt).total_seconds() < QUERY_COOLDOWN_MINUTES * 60:
                        logger.info(
                            "Skipping query '%s' for user %s: cooldown %s minutes",
                            q["query"],
                            user_id,
                            QUERY_COOLDOWN_MINUTES,
                        )
                        continue
                except Exception:
                    pass

        sheet_name = _sheet_name_for_query(q)
        logger.info("Processing query '%s' for user %s", q["query"], user_id)
        if q.get("query_uk"):
            listings = search_olx(
                query=q.get("query_uk"),
                price_from=q.get("price_from"),
                price_to=q.get("price_to"),
                city=q.get("city"),
                photos_only=True,
                pages=SEARCH_PAGES,
                progress_cb=lambda msg, qname=q["query"]: logger.info("%s | %s", qname, msg),
                page_delay=PAGE_DELAY,
                detail_delay=DETAIL_DELAY,
                max_listings=MAX_LISTINGS_PER_QUERY,
            )
            listings_ru = search_olx(
                query=q["query"],
                price_from=q.get("price_from"),
                price_to=q.get("price_to"),
                city=q.get("city"),
                photos_only=True,
                pages=SEARCH_PAGES,
                progress_cb=lambda msg, qname=q["query"]: logger.info("%s | %s", qname, msg),
                page_delay=PAGE_DELAY,
                detail_delay=DETAIL_DELAY,
                max_listings=MAX_LISTINGS_PER_QUERY,
            )
            listings = listings + listings_ru
        else:
            listings = search_olx(
                query=q["query"],
                price_from=q.get("price_from"),
                price_to=q.get("price_to"),
                city=q.get("city"),
                photos_only=True,
                pages=SEARCH_PAGES,
                progress_cb=lambda msg, qname=q["query"]: logger.info("%s | %s", qname, msg),
                page_delay=PAGE_DELAY,
                detail_delay=DETAIL_DELAY,
                max_listings=MAX_LISTINGS_PER_QUERY,
            )
        q["last_run"] = now.isoformat(timespec="seconds")
        logger.info("Query '%s' for user %s returned %s listings", q["query"], user_id, len(listings))
        for l in listings:
            if l.listing_id in seen:
                continue
            seen[l.listing_id] = l.url
            l.sheet_name = sheet_name
            new_listings.append(l)

    added = _append_to_excel(user_id, new_listings)
    if added:
        _prune_excel(_excel_file(user_id), max_age_days=40)
    save_seen_ids(_seen_file(user_id), seen)
    save_queries(_query_file(user_id), queries)

    if added:
        chat_id = _load_chat_id(user_id)
        if chat_id:
            summary = _format_summary(new_listings)
            await app.bot.send_message(chat_id=chat_id, text=summary)

    logger.info("Scrape finished for user %s. Added=%s", user_id, added)
    return added


async def _send_daily_report_for_user(user_id: int, app: Application) -> None:
    chat_id = _load_chat_id(user_id)
    if not chat_id:
        logger.warning("Daily report skipped for user %s: chat_id not set.", user_id)
        return
    excel_path = _excel_file(user_id)
    if not excel_path.exists():
        await app.bot.send_message(chat_id=chat_id, text="Нет данных для отчета.")
        return
    _prune_excel(excel_path, max_age_days=40)
    recent = _load_recent_from_excel(excel_path, hours=24)
    summary = _format_summary(recent, limit=10)
    await app.bot.send_message(chat_id=chat_id, text=summary)
    await app.bot.send_document(chat_id=chat_id, document=excel_path.open("rb"), filename=excel_path.name)


async def _run_scheduled_scrapes(app: Application) -> None:
    users = _current_users()
    for user_id in users:
        await _run_scrape_for_user(user_id, app, force=False)


async def _send_scheduled_reports(app: Application) -> None:
    users = _current_users()
    for user_id in users:
        await _send_daily_report_for_user(user_id, app)


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = _user_from_update(update)
    if user_id is None:
        return
    if not _is_allowed(user_id):
        return
    if update.effective_chat:
        _save_chat_id(user_id, update.effective_chat.id)
    _ensure_user_dir(user_id)
    await update.message.reply_text(HELP_TEXT)


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _user_allowed(update):
        return
    await update.message.reply_text(HELP_TEXT)


async def cmd_myid(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = _user_from_update(update)
    if user_id is None:
        return
    await update.message.reply_text(f"Ваш user_id: {user_id}")


async def cmd_list(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _user_allowed(update):
        return
    user_id = _user_from_update(update)
    if user_id is None:
        return
    queries = load_queries(_query_file(user_id))
    if not queries:
        await update.message.reply_text("Список запросов пуст.")
        return
    lines = ["Ваши запросы:"]
    for i, q in enumerate(queries, start=1):
        city = f", {q['city']}" if q.get("city") else ""
        price = ""
        if q.get("price_from") is not None and q.get("price_to") is not None:
            price = f", от {q['price_from']} до {q['price_to']}"
        extra = f" (uk: {q['query_uk']})" if q.get("query_uk") else ""
        lines.append(f"{i}. {q['query']}{price}{city}{extra}")
    await update.message.reply_text("\n".join(lines))


async def cmd_remove(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _user_allowed(update):
        return
    user_id = _user_from_update(update)
    if user_id is None:
        return
    queries = load_queries(_query_file(user_id))
    if not queries:
        await update.message.reply_text("Список запросов пуст.")
        return
    if not context.args:
        await update.message.reply_text("Укажите номер: /remove 2")
        return
    try:
        idx = int(context.args[0]) - 1
    except ValueError:
        await update.message.reply_text("Номер должен быть числом.")
        return
    if idx < 0 or idx >= len(queries):
        await update.message.reply_text("Нет запроса с таким номером.")
        return
    removed = queries.pop(idx)
    save_queries(_query_file(user_id), queries)
    await update.message.reply_text(f"Удалено: {removed['query']}")


async def cmd_run(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _user_allowed(update):
        return
    user_id = _user_from_update(update)
    if user_id is None:
        return
    await update.message.reply_text("Запускаю сбор...")
    added = await _run_scrape_for_user(user_id, context.application, force=True)
    await update.message.reply_text(f"Готово. Новых объявлений: {added}")


async def cmd_report(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _user_allowed(update):
        return
    user_id = _user_from_update(update)
    if user_id is None:
        return
    await update.message.reply_text("Отправляю файл...")
    await _send_daily_report_for_user(user_id, context.application)


async def cmd_log(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _user_allowed(update):
        return
    if not LOG_FILE.exists():
        await update.message.reply_text("Лог-файл не найден.")
        return
    try:
        lines = LOG_FILE.read_text(encoding="utf-8").splitlines()[-50:]
    except Exception:
        await update.message.reply_text("Не удалось прочитать лог.")
        return
    if not lines:
        await update.message.reply_text("Лог пуст.")
        return
    await update.message.reply_text("\n".join(lines[-50:]))


async def cmd_allow(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_admin_update(update):
        return
    if not context.args:
        await update.message.reply_text("Укажите user_id: /allow 123")
        return
    try:
        user_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("user_id должен быть числом.")
        return
    _add_allowed_user(user_id)
    _ensure_user_dir(user_id)
    await update.message.reply_text(f"Пользователь {user_id} добавлен.")


async def cmd_deny(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_admin_update(update):
        return
    if not context.args:
        await update.message.reply_text("Укажите user_id: /deny 123")
        return
    try:
        user_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("user_id должен быть числом.")
        return
    _remove_allowed_user(user_id)
    await update.message.reply_text(f"Пользователь {user_id} удален.")


async def cmd_users(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_admin_update(update):
        return
    users = _current_users()
    if not users:
        await update.message.reply_text("Список пуст.")
        return
    await update.message.reply_text("Разрешенные пользователи:\n" + "\n".join(str(u) for u in users))


async def on_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = _user_from_update(update)
    if user_id is None:
        return
    if not _is_allowed(user_id):
        return

    if update.effective_chat:
        _save_chat_id(user_id, update.effective_chat.id)

    parsed = _parse_query(update.message.text)
    if not parsed:
        await update.message.reply_text("Не понял формат. Наберите /start для инструкции.")
        return

    queries = load_queries(_query_file(user_id))
    if parsed.get("query"):
        parsed["query_uk"] = _translate_ru_uk(parsed["query"])
    queries.append(parsed)
    save_queries(_query_file(user_id), queries)
    logger.info("Added query for user %s: %s", user_id, parsed)

    city = f", {parsed['city']}" if parsed.get("city") else ""
    price = ""
    if parsed.get("price_from") is not None and parsed.get("price_to") is not None:
        price = f", от {parsed['price_from']} до {parsed['price_to']}"
    await update.message.reply_text(f"Добавлен запрос: {parsed['query']}{price}{city}")


async def cmd_search(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _user_allowed(update):
        return
    user_id = _user_from_update(update)
    if user_id is None:
        return
    if not context.args:
        await update.message.reply_text("Пример: /search Дом, от 100000 до 4000000, г. Одесса")
        return
    raw = " ".join(context.args).strip()
    parsed = _parse_query(raw)
    if not parsed:
        await update.message.reply_text("Не понял формат. Пример: /search Дом, от 100000 до 4000000, г. Одесса")
        return
    parsed["query_uk"] = _translate_ru_uk(parsed["query"]) if parsed.get("query") else None
    queries = load_queries(_query_file(user_id))
    queries.append(parsed)
    save_queries(_query_file(user_id), queries)
    city = f", {parsed['city']}" if parsed.get("city") else ""
    price = ""
    if parsed.get("price_from") is not None and parsed.get("price_to") is not None:
        price = f", от {parsed['price_from']} до {parsed['price_to']}"
    await update.message.reply_text(f"Добавлен запрос: {parsed['query']}{price}{city}")


async def post_init(app: Application) -> None:
    tz = ZoneInfo(TIMEZONE)
    scheduler = AsyncIOScheduler(timezone=tz)

    scheduler.add_job(lambda: asyncio.create_task(_run_scheduled_scrapes(app)), "interval", hours=12)
    scheduler.add_job(lambda: asyncio.create_task(_send_scheduled_reports(app)), "cron", hour=11, minute=0)
    scheduler.start()


def main() -> None:
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        raise SystemExit("TELEGRAM_BOT_TOKEN не задан. Заполните .env")

    ensure_dir(DATA_DIR)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )
    logger.info("Bot starting...")

    app = Application.builder().token(token).post_init(post_init).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("myid", cmd_myid))
    app.add_handler(CommandHandler("list", cmd_list))
    app.add_handler(CommandHandler("remove", cmd_remove))
    app.add_handler(CommandHandler("run", cmd_run))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(CommandHandler("log", cmd_log))
    app.add_handler(CommandHandler("allow", cmd_allow))
    app.add_handler(CommandHandler("deny", cmd_deny))
    app.add_handler(CommandHandler("users", cmd_users))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_message))

    app.run_polling()


if __name__ == "__main__":
    main()
