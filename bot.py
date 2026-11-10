#!/usr/bin/env python3
"""Crew punch-clock Telegram bot.

Employees tap "Punch in" / "Punch out". The bot stores every punch and the
admin gets biweekly hour totals (text + Excel) for payroll.

Setup is via environment variables (see README.md):
    BOT_TOKEN   Telegram bot token from @BotFather (required)
    ADMIN_IDS   Comma-separated Telegram user IDs of admins (required)
    PAY_ANCHOR  First day (a Monday) of a pay period, YYYY-MM-DD.
                Pay periods are 14 days starting from this date.
    TZ          Timezone, default America/Toronto
    DB_PATH     SQLite file, default timeclock.db
    WEBHOOK_URL Public base URL for webhook mode. If unset, the bot polls.
                (Render sets RENDER_EXTERNAL_URL automatically.)
"""

import io
import logging
import os
import sqlite3
import tempfile
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
    Update,
)
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s: %(message)s", level=logging.INFO
)
log = logging.getLogger("timeclock")

# ---------------------------------------------------------------- config ---

TZ = ZoneInfo(os.getenv("TZ", "America/Toronto"))
ADMIN_IDS = {int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()}
PAY_ANCHOR = date.fromisoformat(os.getenv("PAY_ANCHOR", "2026-09-28"))
DB_PATH = os.getenv("DB_PATH", "timeclock.db")
BOT_TOKEN = os.getenv("BOT_TOKEN", "")

ASK_NAME = 1

# ------------------------------------------------------------------- db ---


def db() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute(
        """CREATE TABLE IF NOT EXISTS employees(
               tg_id INTEGER PRIMARY KEY,
               name TEXT NOT NULL,
               approved INTEGER NOT NULL DEFAULT 0,
               active INTEGER NOT NULL DEFAULT 1,
               created_at TEXT NOT NULL)"""
    )
    con.execute(
        """CREATE TABLE IF NOT EXISTS punches(
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               tg_id INTEGER NOT NULL,
               kind TEXT NOT NULL CHECK(kind IN ('IN','OUT')),
               ts TEXT NOT NULL)"""
    )
    con.execute("CREATE INDEX IF NOT EXISTS idx_punches_tg ON punches(tg_id, ts)")
    return con


def now() -> datetime:
    return datetime.now(TZ)


def iso(dt: datetime) -> str:
    return dt.astimezone(TZ).isoformat()


def parse_iso(s: str) -> datetime:
    return datetime.fromisoformat(s).astimezone(TZ)


# ----------------------------------------------------------------- time ---


def period_for(d: date) -> tuple[date, date, int]:
    """Return (period_start, period_end, period_index) for a date.

    Periods are 14 days long starting at PAY_ANCHOR.
    """
    delta = (d - PAY_ANCHOR).days
    idx = delta // 14
    start = PAY_ANCHOR + timedelta(days=14 * idx)
    return start, start + timedelta(days=13), idx


def current_period() -> tuple[date, date, int]:
    return period_for(now().date())


def fmt_hours(seconds: float) -> str:
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return f"{h}h {m:02d}m"


# -------------------------------------------------------------- pairing ---


def pair_punches(rows: list[sqlite3.Row]):
    """Pair IN->OUT punches chronologically.

    Returns (shifts, anomalies) where shifts are (in_dt, out_dt|None) and
    anomalies are human-readable strings (e.g. OUT without IN).
    """
    shifts: list[tuple[datetime, datetime | None]] = []
    anomalies: list[str] = []
    open_in: datetime | None = None
    for r in sorted(rows, key=lambda r: r["ts"]):
        ts = parse_iso(r["ts"])
        if r["kind"] == "IN":
            if open_in is not None:
                anomalies.append(
                    f"Double punch-in on {open_in:%b %d}: "
                    f"{open_in:%H:%M} then {ts:%H:%M} (kept first)"
                )
            else:
                open_in = ts
        else:  # OUT
            if open_in is None:
                anomalies.append(f"Punch-out with no punch-in on {ts:%b %d %H:%M}")
            else:
                shifts.append((open_in, ts))
                open_in = None
    if open_in is not None:
        shifts.append((open_in, None))  # still clocked in
    return shifts, anomalies


def period_totals(start: date, end: date) -> dict:
    """Per-employee totals for a pay period.

    Returns {tg_id: {"name":..., "seconds":..., "days": {date: seconds},
                     "open": bool, "anomalies": [...]}}.
    """
    con = db()
    try:
        emps = {
            r["tg_id"]: r["name"]
            for r in con.execute(
                "SELECT tg_id, name FROM employees WHERE approved=1 AND active=1"
            )
        }
        result: dict = {}
        lo = datetime.combine(start, time.min, TZ).isoformat()
        hi = datetime.combine(end + timedelta(days=1), time.min, TZ).isoformat()
        for tg_id, name in emps.items():
            rows = list(
                con.execute(
                    "SELECT kind, ts FROM punches "
                    "WHERE tg_id=? AND ts>=? AND ts<? ORDER BY ts",
                    (tg_id, lo, hi),
                )
            )
            shifts, anomalies = pair_punches(rows)
            seconds = 0.0
            days: dict[date, float] = {}
            open_shift = False
            for s_in, s_out in shifts:
                if s_out is None:
                    open_shift = True
                    continue
                dur = (s_out - s_in).total_seconds()
                if dur < 0:
                    anomalies.append(f"Negative shift on {s_in:%b %d}, ignored")
                    continue
                if dur > 16 * 3600:
                    anomalies.append(
                        f"Very long shift on {s_in:%b %d} ({fmt_hours(dur)}), check it"
                    )
                seconds += dur
                days[s_in.date()] = days.get(s_in.date(), 0.0) + dur
            result[tg_id] = {
                "name": name,
                "seconds": seconds,
                "days": days,
                "open": open_shift,
                "anomalies": anomalies,
            }
        return result
    finally:
        con.close()


# --------------------------------------------------------------- report ---


def build_report_text(start: date, end: date, totals: dict) -> str:
    lines = [
        f"🕒 Pay period {start:%b %d} – {end:%b %d, %Y}",
        "",
    ]
    grand = 0.0
    for tg_id in sorted(totals, key=lambda t: totals[t]["name"].lower()):
        t = totals[tg_id]
        grand += t["seconds"]
        line = f"• {t['name']}: {fmt_hours(t['seconds'])} ({t['seconds']/3600:.2f}h)"
        if t["open"]:
            line += "  ⚠️ still punched in"
        lines.append(line)
        for a in t["anomalies"]:
            lines.append(f"    ⚠️ {a}")
    lines += ["", f"Total crew hours: {fmt_hours(grand)} ({grand/3600:.2f}h)"]
    return "\n".join(lines)


def build_report_xlsx(start: date, end: date, totals: dict) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Employee", "Total hours (decimal)", "Total (h:mm)"])
    for tg_id in sorted(totals, key=lambda t: totals[t]["name"].lower()):
        t = totals[tg_id]
        dec = round(t["seconds"] / 3600, 2)
        ws.append([t["name"], dec, fmt_hours(t["seconds"])])
    det = wb.create_sheet("Daily detail")
    det.append(["Employee", "Date", "Hours (decimal)", "Hours (h:mm)"])
    for tg_id in sorted(totals, key=lambda t: totals[t]["name"].lower()):
        t = totals[tg_id]
        for d in sorted(t["days"]):
            s = t["days"][d]
            det.append([t["name"], d.isoformat(), round(s / 3600, 2), fmt_hours(s)])
    for sheet in (ws, det):
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = 22
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


async def send_report_to_admins(context: ContextTypes.DEFAULT_TYPE, periods_ago: int = 0):
    start, end, _ = current_period()
    if periods_ago:
        start -= timedelta(days=14 * periods_ago)
        end -= timedelta(days=14 * periods_ago)
    totals = period_totals(start, end)
    text = build_report_text(start, end, totals)
    xlsx = build_report_xlsx(start, end, totals)
    fname = f"hours_{start:%Y-%m-%d}_{end:%Y-%m-%d}.xlsx"
    for admin_id in ADMIN_IDS:
        try:
            await context.bot.send_message(admin_id, text)
            await context.bot.send_document(
                admin_id, document=xlsx, filename=fname,
                caption=f"Hours {start:%b %d} – {end:%b %d}",
            )
            xlsx.seek(0)
        except Exception:  # noqa: BLE001 - one unreachable admin must not break others
            log.exception("Could not send report to admin %s", admin_id)


# -------------------------------------------------------------- keyboards ---


def is_punched_in(tg_id: int) -> datetime | None:
    con = db()
    try:
        rows = list(
            con.execute(
                "SELECT kind, ts FROM punches WHERE tg_id=? ORDER BY ts", (tg_id,)
            )
        )
    finally:
        con.close()
    shifts, _ = pair_punches(rows)
    if shifts and shifts[-1][1] is None:
        return shifts[-1][0]
    return None


def main_keyboard(tg_id: int) -> ReplyKeyboardMarkup:
    punched = is_punched_in(tg_id)
    if punched:
        rows = [[KeyboardButton("🔴 Punch out")]]
    else:
        rows = [[KeyboardButton("🟢 Punch in")]]
    rows.append([KeyboardButton("📊 My hours")])
    if tg_id in ADMIN_IDS:
        rows.append([KeyboardButton("📋 Biweekly report")])
        rows.append([KeyboardButton("💾 Backup data")])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


# --------------------------------------------------------------- handlers ---


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    con = db()
    try:
        row = con.execute(
            "SELECT name, approved FROM employees WHERE tg_id=?", (user.id,)
        ).fetchone()
    finally:
        con.close()

    if user.id in ADMIN_IDS:
        await update.message.reply_text(
            "You're registered as admin. Employees /start here to register; "
            "you'll get an approve button.\n\n"
            "Commands:\n"
            "/report [n] – hours for the pay period n periods ago (0 = current)\n"
            "/employees – list crew\n"
            "/manual Name YYYY-MM-DD HH:MM-HH:MM – add a missed shift\n"
            "/dellast Name – delete an employee's most recent punch\n"
            "/remove Name – deactivate an employee\n"
            "/backup – get a full copy of all data (tap before any bot update)",
            reply_markup=main_keyboard(user.id),
        )
        return ConversationHandler.END

    if row and row["approved"]:
        await update.message.reply_text(
            f"Welcome back, {row['name']}!",
            reply_markup=main_keyboard(user.id),
        )
        return ConversationHandler.END
    if row and not row["approved"]:
        await update.message.reply_text(
            "Your registration is waiting for approval. Hang tight!"
        )
        return ConversationHandler.END

    await update.message.reply_text(
        "Welcome! What's your full name? (as it should appear on hours reports)"
    )
    return ASK_NAME


async def got_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    name = update.message.text.strip()
    con = db()
    try:
        con.execute(
            "INSERT OR REPLACE INTO employees(tg_id, name, approved, active, created_at)"
            " VALUES(?, ?, 0, 1, ?)",
            (user.id, name, iso(now())),
        )
        con.commit()
    finally:
        con.close()

    kb = InlineKeyboardMarkup(
        [[
            InlineKeyboardButton("✅ Approve", callback_data=f"approve:{user.id}"),
            InlineKeyboardButton("❌ Deny", callback_data=f"deny:{user.id}"),
        ]]
    )
    for admin_id in ADMIN_IDS:
        try:
            await context.bot.send_message(
                admin_id,
                f"🙋 New registration: *{name}* (id {user.id}). Approve?",
                parse_mode="Markdown",
                reply_markup=kb,
            )
        except Exception:
            log.exception("Could not notify admin %s", admin_id)
    await update.message.reply_text(
        "Thanks! Your boss has been asked to approve you. "
        "You'll get a message once you're in."
    )
    return ConversationHandler.END


async def approval_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    action, tg_id_s = query.data.split(":")
    tg_id = int(tg_id_s)
    con = db()
    try:
        if action == "approve":
            con.execute("UPDATE employees SET approved=1 WHERE tg_id=?", (tg_id,))
            con.commit()
            row = con.execute(
                "SELECT name FROM employees WHERE tg_id=?", (tg_id,)
            ).fetchone()
            name = row["name"] if row else "there"
            await query.edit_message_text(f"✅ Approved {name}.")
            try:
                await context.bot.send_message(
                    tg_id,
                    f"You're approved, {name}! Use the buttons below to punch in/out.",
                    reply_markup=main_keyboard(tg_id),
                )
            except Exception:
                log.exception("Could not notify approved employee %s", tg_id)
        else:
            con.execute("DELETE FROM employees WHERE tg_id=?", (tg_id,))
            con.commit()
            await query.edit_message_text("❌ Denied and removed.")
            try:
                await context.bot.send_message(
                    tg_id, "Your registration wasn't approved. Talk to your boss."
                )
            except Exception:
                log.exception("Could not notify denied employee %s", tg_id)
    finally:
        con.close()


def record_punch(tg_id: int, kind: str) -> datetime:
    ts = now()
    con = db()
    try:
        con.execute(
            "INSERT INTO punches(tg_id, kind, ts) VALUES(?, ?, ?)",
            (tg_id, kind, iso(ts)),
        )
        con.commit()
    finally:
        con.close()
    return ts


def require_approved(update: Update) -> tuple[bool, str]:
    """Returns (ok, name)."""
    con = db()
    try:
        row = con.execute(
            "SELECT name FROM employees WHERE tg_id=? AND approved=1 AND active=1",
            (update.effective_user.id,),
        ).fetchone()
    finally:
        con.close()
    if not row:
        return False, ""
    return True, row["name"]


async def punch_in(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ok, name = require_approved(update)
    if not ok:
        await update.message.reply_text("Please /start and get approved first.")
        return
    already = is_punched_in(update.effective_user.id)
    if already:
        await update.message.reply_text(
            f"You're already punched in (since {already:%H:%M}).",
            reply_markup=main_keyboard(update.effective_user.id),
        )
        return
    ts = record_punch(update.effective_user.id, "IN")
    await update.message.reply_text(
        f"🟢 {name} punched in at {ts:%H:%M}. Have a good shift!",
        reply_markup=main_keyboard(update.effective_user.id),
    )


async def punch_out(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ok, name = require_approved(update)
    if not ok:
        await update.message.reply_text("Please /start and get approved first.")
        return
    since = is_punched_in(update.effective_user.id)
    if since is None:
        await update.message.reply_text(
            "You're not punched in right now.",
            reply_markup=main_keyboard(update.effective_user.id),
        )
        return
    ts = record_punch(update.effective_user.id, "OUT")
    dur = (ts - since).total_seconds()
    await update.message.reply_text(
        f"🔴 {name} punched out at {ts:%H:%M}.\nShift: {fmt_hours(dur)}",
        reply_markup=main_keyboard(update.effective_user.id),
    )


async def my_hours(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ok, name = require_approved(update)
    if not ok:
        await update.message.reply_text("Please /start and get approved first.")
        return
    start, end, _ = current_period()
    totals = period_totals(start, end)
    t = totals.get(update.effective_user.id)
    if not t or (t["seconds"] == 0 and not t["open"]):
        await update.message.reply_text(
            f"No hours yet for {start:%b %d} – {end:%b %d}."
        )
        return
    lines = [f"📊 {name} — {start:%b %d} to {end:%b %d}"]
    for d in sorted(t["days"]):
        lines.append(f"{d:%a %b %d}: {fmt_hours(t['days'][d])}")
    if t["open"]:
        lines.append("⚠️ You're still punched in.")
    lines.append(f"Total: {fmt_hours(t['seconds'])} ({t['seconds']/3600:.2f}h)")
    await update.message.reply_text("\n".join(lines))


async def admin_report_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    n = 0
    if context.args:
        try:
            n = int(context.args[0])
        except ValueError:
            pass
    await send_report_to_admins(context, periods_ago=n)
    if n == 0:
        await update.message.reply_text("Report sent above. 👆")


async def report_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    await send_report_to_admins(context)


async def backup_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin only: send a full snapshot of the database as a file.

    The snapshot uses SQLite's backup API, so it is consistent even if
    someone punches at the same moment. Save the file somewhere safe:
    before any bot update, tap Backup first; if the hosted copy is ever
    lost, this file is what restores every employee and punch.
    """
    if update.effective_user.id not in ADMIN_IDS:
        return
    stamp = now().strftime("%Y-%m-%d-%H%M")
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            delete=False, suffix=".db", prefix="timeclock-backup-"
        ) as tmp:
            tmp_path = tmp.name
        src = db()
        try:
            dst = sqlite3.connect(tmp_path)
            try:
                src.backup(dst)
            finally:
                dst.close()
        finally:
            src.close()
        with open(tmp_path, "rb") as fh:
            await update.message.reply_document(
                document=fh,
                filename=f"timeclock-backup-{stamp}.db",
                caption=(
                    "Full backup of all employees and punches. "
                    "Save this file. If the bot's data is ever lost, "
                    "send this file to Muse to restore it."
                ),
            )
    except Exception:
        log.exception("Backup failed")
        await update.message.reply_text("Backup failed. Try again in a minute.")
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)


async def employees_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    con = db()
    try:
        rows = list(
            con.execute(
                "SELECT name, approved, active FROM employees ORDER BY name"
            )
        )
    finally:
        con.close()
    if not rows:
        await update.message.reply_text("No employees yet.")
        return
    lines = []
    for r in rows:
        status = "✅" if (r["approved"] and r["active"]) else "⏳"
        lines.append(f"{status} {r['name']}")
    await update.message.reply_text("\n".join(lines))


async def manual_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """ /manual Name YYYY-MM-DD HH:MM-HH:MM  — add a missed shift. """
    if update.effective_user.id not in ADMIN_IDS:
        return
    try:
        name = context.args[0]
        day = date.fromisoformat(context.args[1])
        in_s, out_s = context.args[2].split("-")
        t_in = datetime.combine(
            day, datetime.strptime(in_s, "%H:%M").time(), TZ
        )
        t_out = datetime.combine(
            day, datetime.strptime(out_s, "%H:%M").time(), TZ
        )
    except (IndexError, ValueError):
        await update.message.reply_text(
            "Usage: /manual Name 2026-10-02 08:00-16:30"
        )
        return
    con = db()
    try:
        row = con.execute(
            "SELECT tg_id FROM employees WHERE lower(name)=lower(?) AND active=1",
            (name,),
        ).fetchone()
        if not row:
            await update.message.reply_text(f"No active employee named '{name}'.")
            return
        con.execute(
            "INSERT INTO punches(tg_id, kind, ts) VALUES(?, 'IN', ?)",
            (row["tg_id"], iso(t_in)),
        )
        con.execute(
            "INSERT INTO punches(tg_id, kind, ts) VALUES(?, 'OUT', ?)",
            (row["tg_id"], iso(t_out)),
        )
        con.commit()
    finally:
        con.close()
    await update.message.reply_text(
        f"Added shift for {name}: {day:%b %d} {in_s}–{out_s}."
    )


async def dellast_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """ /dellast Name — delete an employee's most recent punch. """
    if update.effective_user.id not in ADMIN_IDS:
        return
    name = " ".join(context.args).strip()
    if not name:
        await update.message.reply_text("Usage: /dellast Name")
        return
    con = db()
    try:
        row = con.execute(
            "SELECT tg_id FROM employees WHERE lower(name)=lower(?) AND active=1",
            (name,),
        ).fetchone()
        if not row:
            await update.message.reply_text(f"No active employee named '{name}'.")
            return
        last = con.execute(
            "SELECT id, kind, ts FROM punches WHERE tg_id=? ORDER BY ts DESC LIMIT 1",
            (row["tg_id"],),
        ).fetchone()
        if not last:
            await update.message.reply_text(f"{name} has no punches.")
            return
        con.execute("DELETE FROM punches WHERE id=?", (last["id"],))
        con.commit()
        ts = parse_iso(last["ts"])
        await update.message.reply_text(
            f"Deleted {name}'s last punch: {last['kind']} at {ts:%b %d %H:%M}."
        )
    finally:
        con.close()


async def remove_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """ /remove Name — deactivate an employee (keeps history). """
    if update.effective_user.id not in ADMIN_IDS:
        return
    name = " ".join(context.args).strip()
    if not name:
        await update.message.reply_text("Usage: /remove Name")
        return
    con = db()
    try:
        cur = con.execute(
            "UPDATE employees SET active=0 WHERE lower(name)=lower(?)", (name,)
        )
        con.commit()
        if cur.rowcount:
            await update.message.reply_text(f"{name} deactivated (history kept).")
        else:
            await update.message.reply_text(f"No employee named '{name}'.")
    finally:
        con.close()


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Tap 🟢 Punch in when you start work and 🔴 Punch out when you finish.\n"
        "📊 My hours shows your time for the current 2-week pay period."
    )


# ------------------------------------------------------------ auto report ---


async def friday_report(context: ContextTypes.DEFAULT_TYPE):
    """Runs every Friday; sends the report only on pay-period end Fridays."""
    today = now().date()
    start, end, _ = current_period()
    if today == end and today.weekday() == 4:
        await send_report_to_admins(context)


# ------------------------------------------------------------------- main ---


def main() -> None:
    if not BOT_TOKEN:
        raise SystemExit("BOT_TOKEN env var is required.")
    if not ADMIN_IDS:
        raise SystemExit("ADMIN_IDS env var is required (your Telegram user ID).")

    db().close()  # create tables

    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={ASK_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_name)]},
        fallbacks=[CommandHandler("start", start)],
    )
    app.add_handler(conv)
    app.add_handler(CallbackQueryHandler(approval_button, pattern="^(approve|deny):"))
    app.add_handler(MessageHandler(filters.Regex("^🟢 Punch in$"), punch_in))
    app.add_handler(MessageHandler(filters.Regex("^🔴 Punch out$"), punch_out))
    app.add_handler(MessageHandler(filters.Regex("^📊 My hours$"), my_hours))
    app.add_handler(MessageHandler(filters.Regex("^📋 Biweekly report$"), report_button))
    app.add_handler(MessageHandler(filters.Regex("^💾 Backup data$"), backup_cmd))
    app.add_handler(CommandHandler("backup", backup_cmd))
    app.add_handler(CommandHandler("report", admin_report_cmd))
    app.add_handler(CommandHandler("employees", employees_cmd))
    app.add_handler(CommandHandler("manual", manual_cmd))
    app.add_handler(CommandHandler("dellast", dellast_cmd))
    app.add_handler(CommandHandler("remove", remove_cmd))
    app.add_handler(CommandHandler("help", help_cmd))

    # Auto-send the report every Friday at 6pm; the job itself checks
    # whether that Friday is the end of a pay period.
    app.job_queue.run_daily(friday_report, time=time(18, 0, tzinfo=TZ), days=(4,))

    webhook_url = os.getenv("WEBHOOK_URL") or os.getenv("RENDER_EXTERNAL_URL")
    if webhook_url:
        port = int(os.getenv("PORT", "10000"))
        log.info("Starting in webhook mode: %s", webhook_url)
        app.run_webhook(
            listen="0.0.0.0",
            port=port,
            url_path="telegram",
            webhook_url=f"{webhook_url.rstrip('/')}/telegram",
        )
    else:
        log.info("Starting in polling mode")
        app.run_polling()


if __name__ == "__main__":
    main()
