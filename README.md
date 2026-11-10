# Crew Timeclock — Telegram punch-in/out bot

Your crew taps **🟢 Punch in** / **🔴 Punch out** in Telegram. Every two
weeks you automatically get everyone's hours (message + Excel file) for payroll.

## What you need (about 10 minutes)

### 1. Create the bot (2 min, in Telegram)
1. Open Telegram and message **@BotFather**.
2. Send `/newbot`, pick a name (e.g. `Alex Landscaping Clock`) and a username
   ending in `bot` (e.g. `alex_landscaping_clock_bot`).
3. BotFather replies with a **token** — long string like `123456:ABC-DEF...`.
   Copy it. (Keep it private; it's the bot's password.)

### 2. Get your Telegram user ID (1 min)
Message **@userinfobot** in Telegram — it replies with your numeric ID.
You'll paste this as `ADMIN_IDS` so the bot knows you're the boss.

### 3. Deploy it free on Render (5 min)
1. Put this folder on GitHub (or use Render's "deploy from local" via
   `render.yaml` blueprint).
2. In [Render](https://render.com) → **New → Blueprint**, point it at the repo.
3. When asked for environment variables, fill in:
   - `BOT_TOKEN` → the token from BotFather
   - `ADMIN_IDS` → your numeric Telegram ID
   - `PAY_ANCHOR` → the Monday your current 2-week pay period started
     (default `2026-09-28`; change it to match your real payroll cycle)
   - `TZ` → `America/Toronto` (already set)
4. Deploy. Render gives the service a public URL; the bot picks it up
   automatically (`RENDER_EXTERNAL_URL`) and runs in webhook mode.

### 4. Register the crew
1. You open the bot in Telegram and send `/start`.
2. Each employee opens the bot, sends `/start`, types their full name.
3. You get an **Approve / Deny** button for each one. Done.

## Daily use

- **Employees:** tap 🟢 Punch in at the start of the day, 🔴 Punch out at the
  end. 📊 My hours shows their time for the current pay period.
- **You:** tap 📋 Biweekly report any time, or just wait — the bot sends the
  full report (message + `.xlsx` file) automatically every second Friday at
  6 PM.

## Fixing mistakes

- `/manual Name 2026-10-02 08:00-16:30` — add a shift someone forgot to punch.
- `/dellast Name` — delete an employee's most recent punch (undo a mis-tap).
- `/report 1` — hours for the previous pay period (`/report 2` for the one
  before that, etc.).
- `/employees` — list the crew. `/remove Name` — deactivate someone who left
  (their history is kept).

The report flags problems for you: still-punched-in shifts, punch-outs with
no punch-in, double punch-ins, and shifts over 16 hours.

## Running it elsewhere

Anywhere Python runs works. Without `WEBHOOK_URL` set, the bot uses polling
(good for a home server or VPS):

```bash
pip install -r requirements.txt
BOT_TOKEN=... ADMIN_IDS=... PAY_ANCHOR=2026-09-28 python bot.py
```

Data lives in a single SQLite file (`timeclock.db`) next to the bot.
