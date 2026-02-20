from __future__ import annotations

import asyncio
import hashlib
import logging
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from dotenv import load_dotenv
from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters
from zoneinfo import ZoneInfo

from olx_scraper import Listing, search_olx
from storage import ensure_dir, load_queries, load_seen_ids, save_queries, save_seen_ids

load_dotenv()

DATA_DIR = Path(os.getenv("DATA_DIR", "data"))
EXCEL_FILE = Path(os.getenv("EXCEL_FILE", "data/olx.xlsx"))
QUERY_FILE = Path(os.getenv("QUERY_FILE", "data/queries.json"))
SEEN_FILE = Path(os.getenv("SEEN_FILE", "data/seen_ids.json"))
TIMEZONE = os.getenv("TIMEZONE", "Europe/Kyiv")
CHAT_ID_ENV = os.getenv("TELEGRAM_CHAT_ID")
SEARCH_PAGES = int(os.getenv("SEARCH_PAGES", "3"))
PAGE_DELAY = float(os.getenv("PAGE_DELAY", "0"))
DETAIL_DELAY = float(os.getenv("DETAIL_DELAY", "0"))
MAX_LISTINGS_PER_QUERY = os.getenv("MAX_LISTINGS_PER_QUERY")
MAX_LISTINGS_PER_QUERY = int(MAX_LISTINGS_PER_QUERY) if MAX_LISTINGS_PER_QUERY else None
QUERY_COOLDOWN_MINUTES = int(os.getenv("QUERY_COOLDOWN_MINUTES", "30"))
LOG_FILE = Path(os.getenv("LOG_FILE", "data/bot.log"))

logger = logging.getLogger("olx-bot")

HELP_TEXT = (
    "Отправьте запрос в формате:\n"
    "`Название, от 1000 до 50000, Город`\n"
    "или без города:\n"
    "`Название, от 1000 до 50000`\n\n"
    "Команды:\n"
    "/list — список запросов\n"
    "/remove N — удалить запрос по номеру\n"
    "/run — выполнить сбор сейчас\n"
    "/report — отправить Excel сейчас\n"
    "/log — последние строки лога"
)


def _load_chat_id() -> Optional[int]:
    if CHAT_ID_ENV:
        return int(CHAT_ID_ENV)
    chat_file = DATA_DIR / "chat_id.txt"
    if chat_file.exists():
        return int(chat_file.read_text(encoding="utf-8").strip())
    return None


def _save_chat_id(chat_id: int) -> None:
    ensure_dir(DATA_DIR)
    chat_file = DATA_DIR / "chat_id.txt"
    chat_file.write_text(str(chat_id), encoding="utf-8")


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

    # Example: "от 1000 до 50000"
    match = re.search(r"от\s*(\d+)\s*до\s*(\d+)", price_text, re.IGNORECASE)
    if match:
        price_from = int(match.group(1))
        price_to = int(match.group(2))
        if len(parts) >= 3:
            city = parts[2]
    else:
        # If city is second part
        if len(parts) >= 2:
            city = parts[1]

    return {
        "query": query,
        "price_from": price_from,
        "price_to": price_to,
        "city": city,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }


def _append_to_excel(listings: List[Listing]) -> int:
    if not listings:
        return 0

    ensure_dir(DATA_DIR)

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

    if EXCEL_FILE.exists():
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            for sheet_name, rows in rows_by_sheet.items():
                df = pd.DataFrame(rows)
                sheet = writer.sheets.get(sheet_name)
                start_row = sheet.max_row if sheet else 0
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=start_row == 0, startrow=start_row)
    else:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            for sheet_name, rows in rows_by_sheet.items():
                df = pd.DataFrame(rows)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    _format_excel()
    return sum(len(r) for r in rows_by_sheet.values())


def _prune_excel(max_age_days: int = 40) -> int:
    if not EXCEL_FILE.exists():
        return 0
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
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
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        for name, df in sheets_out.items():
            df.to_excel(writer, sheet_name=name, index=False)
    _format_excel()
    return total_removed


def _format_summary(listings: List[Listing], limit: int = 5) -> str:
    if not listings:
        return "Новых объявлений: 0"
    lines = [f"Новых объявлений: {len(listings)}", "Топ:"]
    for l in listings[:limit]:
        price = f"{l.price} {l.currency}" if l.price else "Цена не указана"
        location = l.location or "Локация не указана"
        lines.append(f"- {l.title} | {price} | {location}")
    return "\n".join(lines)


def _load_recent_from_excel(hours: int = 24) -> List[Listing]:
    if not EXCEL_FILE.exists():
        return []
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
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
    cleaned = re.sub(r"[\\\\/:*?\\[\\]]", " ", base).strip()
    if len(cleaned) <= 31:
        return cleaned or "Sheet"
    digest = hashlib.md5(base.encode("utf-8")).hexdigest()[:6]
    trimmed = cleaned[:24].rstrip()
    return f"{trimmed}~{digest}"


def _format_excel() -> None:
    if not EXCEL_FILE.exists():
        return
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font
    except Exception:
        return
    wb = load_workbook(EXCEL_FILE)
    header_font = Font(bold=True)
    wrap = Alignment(wrap_text=True, vertical="top")
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = header_font
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = wrap
        # Auto column widths
        for col_cells in ws.columns:
            max_len = 0
            col = col_cells[0].column_letter
            for cell in col_cells[:500]:
                if cell.value is None:
                    continue
                max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col].width = min(max(10, max_len + 2), 60)
    wb.save(EXCEL_FILE)


async def _run_scrape_job(app: Application, force: bool = False) -> int:
    queries = load_queries(QUERY_FILE)
    if not queries:
        logger.info("No queries found; skipping scrape.")
        return 0

    seen = load_seen_ids(SEEN_FILE)
    new_listings: List[Listing] = []
    now = datetime.now()

    for q in queries:
        if not force:
            last_run = q.get("last_run")
            if last_run:
                try:
                    last_dt = datetime.fromisoformat(last_run)
                    if (now - last_dt).total_seconds() < QUERY_COOLDOWN_MINUTES * 60:
                        logger.info("Skipping query '%s': cooldown %s minutes", q["query"], QUERY_COOLDOWN_MINUTES)
                        continue
                except Exception:
                    pass
        sheet_name = _sheet_name_for_query(q)
        logger.info("Processing query: %s", q["query"])
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
        logger.info("Query '%s' returned %s listings", q["query"], len(listings))
        for l in listings:
            if l.listing_id in seen:
                continue
            seen[l.listing_id] = l.url
            l.sheet_name = sheet_name
            new_listings.append(l)

    added = _append_to_excel(new_listings)
    if added:
        _prune_excel(max_age_days=40)
    save_seen_ids(SEEN_FILE, seen)
    save_queries(QUERY_FILE, queries)

    if added:
        chat_id = _load_chat_id()
        if chat_id:
            summary = _format_summary(new_listings)
            await app.bot.send_message(chat_id=chat_id, text=summary)
    logger.info("Scrape finished. Added=%s", added)
    return added


async def _send_daily_report(app: Application) -> None:
    chat_id = _load_chat_id()
    if not chat_id:
        logger.warning("Daily report skipped: chat_id not set.")
        return
    if not EXCEL_FILE.exists():
        await app.bot.send_message(chat_id=chat_id, text="Нет данных для отчета.")
        return
    _prune_excel(max_age_days=40)
    recent = _load_recent_from_excel(hours=24)
    summary = _format_summary(recent, limit=10)
    await app.bot.send_message(chat_id=chat_id, text=summary)
    await app.bot.send_document(chat_id=chat_id, document=EXCEL_FILE.open("rb"), filename=EXCEL_FILE.name)


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_chat:
        _save_chat_id(update.effective_chat.id)
    await update.message.reply_text(HELP_TEXT, parse_mode=ParseMode.MARKDOWN)


async def cmd_list(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    queries = load_queries(QUERY_FILE)
    if not queries:
        await update.message.reply_text("Список запросов пуст.")
        return
    lines = ["Ваши запросы:"]
    for i, q in enumerate(queries, start=1):
        city = f", {q['city']}" if q.get("city") else ""
        price = ""
        if q.get("price_from") is not None and q.get("price_to") is not None:
            price = f", от {q['price_from']} до {q['price_to']}"
        lines.append(f"{i}. {q['query']}{price}{city}")
    await update.message.reply_text("\n".join(lines))


async def cmd_remove(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    queries = load_queries(QUERY_FILE)
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
    save_queries(QUERY_FILE, queries)
    await update.message.reply_text(f"Удалено: {removed['query']}")


async def cmd_run(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text("Запускаю сбор...")
    added = await _run_scrape_job(context.application, force=True)
    await update.message.reply_text(f"Готово. Новых объявлений: {added}")


async def cmd_report(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text("Отправляю файл...")
    await _send_daily_report(context.application)


async def cmd_log(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
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


async def on_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    if update.effective_chat:
        _save_chat_id(update.effective_chat.id)

    parsed = _parse_query(update.message.text)
    if not parsed:
        await update.message.reply_text("Не понял формат. Наберите /start для инструкции.")
        return

    queries = load_queries(QUERY_FILE)
    queries.append(parsed)
    save_queries(QUERY_FILE, queries)
    logger.info("Added query: %s", parsed)

    city = f", {parsed['city']}" if parsed.get("city") else ""
    price = ""
    if parsed.get("price_from") is not None and parsed.get("price_to") is not None:
        price = f", от {parsed['price_from']} до {parsed['price_to']}"
    await update.message.reply_text(f"Добавлен запрос: {parsed['query']}{price}{city}")


async def post_init(app: Application) -> None:
    tz = ZoneInfo(TIMEZONE)
    scheduler = AsyncIOScheduler(timezone=tz)

    scheduler.add_job(lambda: asyncio.create_task(_run_scrape_job(app)), "interval", hours=12)
    scheduler.add_job(lambda: asyncio.create_task(_send_daily_report(app)), "cron", hour=11, minute=0)
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
    app.add_handler(CommandHandler("list", cmd_list))
    app.add_handler(CommandHandler("remove", cmd_remove))
    app.add_handler(CommandHandler("run", cmd_run))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(CommandHandler("log", cmd_log))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_message))

    app.run_polling()


if __name__ == "__main__":
    main()
