# Multi-User Access Control + Per-User Data (Branch: `multi_users`)

## Summary
Implement a multi-user version of the OLX bot where:
- Only admin-approved users can use the bot.
- Each user has separate queries and Excel output stored under `data/users/<user_id>/`.
- Admin manages allowed users via Telegram commands.
- A new git branch `multi_users` will contain all changes.

## Scope
- Access control (allowlist) enforced on all commands and messages.
- Per-user storage for `queries.json`, `seen_ids.json`, `olx.xlsx`, and `chat_id.txt`.
- Admin-only Telegram commands to add/remove/list users.
- A `/myid` command so users can send their Telegram user ID to admin.

## Key Decisions (Locked In)
- Admins defined by `ADMIN_CHAT_IDS` in `.env` (comma-separated numeric IDs).
- Data separated by **Telegram user ID**.
- No automatic migration of existing single-user data; old global files remain untouched.

---

## Data Layout
```
data/
  allowed_users.json         # allowlist (list of user IDs)
  users/
    <user_id>/
      queries.json
      seen_ids.json
      olx.xlsx
      chat_id.txt            # equals user_id for private chat; still stored
      bot.log (optional, if per-user logs are desired later)
```

## New/Updated Environment Variables
- `ADMIN_CHAT_IDS=123456,7891011`
- (optional) `ALLOWED_USERS_FILE=data/allowed_users.json` (default)
- Existing env vars remain global (e.g., `SEARCH_PAGES`, delays, etc.).

---

## Bot Command Changes
### Admin-only
- `/allow <user_id>`: add to allowlist
- `/deny <user_id>`: remove from allowlist
- `/users`: list allowed user IDs
- `/myid`: returns caller’s `user_id` (for non-admins too)

### User Commands (Allowed Users Only)
- `/start`, `/list`, `/remove`, `/run`, `/report`, `/log`
- User queries saved under their own folder.

### Unauthorized Users
- **Silent ignore** (no responses).

---

## Core Refactor Plan
### 1) Storage Layer
- Add helpers to resolve per-user paths:
  - `get_user_dir(user_id)`
  - `get_query_file(user_id)`
  - `get_seen_file(user_id)`
  - `get_excel_file(user_id)`
  - `get_chat_id_file(user_id)`
- Allowlist helpers:
  - `load_allowed_users(path)`
  - `save_allowed_users(path)`

### 2) Access Control Middleware
- Implement `is_admin(user_id)` based on `ADMIN_CHAT_IDS`.
- Implement `is_allowed(user_id)` based on allowlist.
- Apply guard in every handler:
  - If not allowed and not admin -> silently return.
  - Admins always allowed even if not in allowlist.

### 3) Per-User Job Execution
Refactor `_run_scrape_job` to:
- `run_scrape_for_user(user_id, force=False)`
- Load that user’s queries and seen_ids.
- Write output to that user’s Excel and prune for that user.
- Save last_run per-query in that user’s queries file.

Add wrappers:
- Scheduled job loops through allowed users and runs each.
- Daily report loops through allowed users and sends each their file.

### 4) Per-User Chat ID
- On `/start` or any message, save `chat_id` to the user’s folder.
- `chat_id` used for sending reports and summaries.

---

## Excel Formatting
Keep current formatting logic, but operate on each user’s file independently.

---

## Commands and Behaviors
### `/allow <user_id>`
- Admin only.
- Adds to allowlist.
- Optionally creates the user folder.

### `/deny <user_id>`
- Admin only.
- Removes from allowlist.
- Does **not** delete user data (safe by default).

### `/users`
- Admin only.
- Returns list of allowed user IDs.

### `/myid`
- Anyone: returns their `user_id` and brief instructions.

---

## Edge Cases & Failure Modes
- User has not started bot yet but is allowed:
  - Daily report skipped for that user until `chat_id.txt` exists.
- User has no queries:
  - Log and skip their scrape.
- Cooldown:
  - Only applies per user per query.
  - `/run` forces bypass cooldown **for the calling user only**.
- Telegram Conflict:
  - Still one bot process; systemd handles restart.

---

## Tests / Manual Verification
1) **Admin allow/deny**
   - As admin: `/allow <id>` updates allowlist and responds.
   - `/users` lists IDs.
   - `/deny <id>` removes.

2) **Unauthorized user**
   - Send any message -> no response.

3) **Allowed user**
   - `/start` creates `data/users/<id>/`.
   - Send query -> stored in that folder.
   - `/run` -> Excel updated only in that folder.

4) **Daily report**
   - Allowed user receives their own file.
   - Other users unaffected.

5) **Cooldown**
   - Run twice within 30 min -> second is skipped (unless `/run`).

---

## Branching / Workflow
1) Create branch `multi_users`.
2) Implement refactor + commands + allowlist.
3) Update README and `.env.example` to document admin and multi-user usage.
4) Restart service after deploy.

---

## Assumptions
- Users interact in **private chats**, so `chat_id == user_id`.
- Old global data remains in place but will no longer be used; manual migration can be done later if desired.
- All global settings (`SEARCH_PAGES`, delays, etc.) remain shared by all users.

If you want “soft migration” for the admin (copy existing data into admin folder), we can add that as an optional admin command later (e.g., `/migrate_legacy`).
