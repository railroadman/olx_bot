# OLX Telegram бот

Бот получает запросы из Telegram, собирает объявления с `olx.ua`, сохраняет в Excel и отправляет ежедневный отчет. Поддерживается multi-user режим с индивидуальными запросами и файлами для каждого пользователя.

## Установка

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Заполните `.env`:
- `TELEGRAM_BOT_TOKEN`
- `ADMIN_CHAT_IDS` (список admin user_id через запятую)

## Multi-user структура

```
data/
  allowed_users.json
  users/
    <user_id>/
      queries.json
      seen_ids.json
      olx.xlsx
      chat_id.txt
```

## Команды

### Пользователь
- `/start` — помощь
- `/myid` — показать ваш user_id
- `/list` — список запросов
- `/remove N` — удалить запрос по номеру
- `/run` — выполнить сбор сейчас (форсирует cooldown)
- `/report` — отправить Excel сейчас
- `/log` — последние строки лога

### Админ
- `/allow <user_id>` — добавить пользователя
- `/deny <user_id>` — удалить пользователя
- `/users` — список разрешенных пользователей

## Формат запросов

- `Название, от 1000 до 50000, Город`
- `Название, от 1000 до 50000`

## Запуск

```bash
python3 bot.py
```

## Примечания

- Фильтр объявлений: не старше ~92 дней.
- Дедупликация по `listing_id` из URL.
- Данные каждого пользователя хранятся отдельно.
