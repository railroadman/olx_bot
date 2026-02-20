# OLX Telegram бот

Бот получает запросы из Telegram, собирает объявления с `olx.ua` каждые 4 часа, сохраняет в Excel и отправляет ежедневный отчет в 11:00 по Киеву.

## Установка

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Заполните `.env`:
- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID` (необязательно — бот сохранит chat_id после команды `/start`)

## Получение chat_id

1. Напишите боту `/start`.
2. Выполните на сервере:

```bash
curl -s "https://api.telegram.org/bot<YOUR_TOKEN>/getUpdates" | python3 -m json.tool
```

В ответе найдите `chat.id`.

## Формат запросов

В Telegram отправьте сообщение:
- `Название, от 1000 до 50000, Город`
- `Название, от 1000 до 50000`

Команды:
- `/list` — список запросов
- `/remove N` — удалить запрос по номеру
- `/run` — выполнить сбор сейчас

## Запуск

```bash
python3 bot.py
```

## Примечания

- Фильтр объявлений: не старше ~92 дней.
- Дедупликация по `listing_id` из URL.
- Данные сохраняются в `data/olx.xlsx`.
- Если OLX поменяет структуру HTML, потребуется обновить селекторы в `olx_scraper.py`.
