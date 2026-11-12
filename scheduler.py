"""Create reminder rows for approved events and send the ones that are due."""

import logging
import threading
from datetime import datetime, timedelta

import db
import mailer
from reminder_logic import Offset, event_start, parse_offsets, parse_recipients, planned_reminders, render_email

logger = logging.getLogger(__name__)

MAX_ATTEMPTS = 3
RETRY_AFTER = timedelta(minutes=5)
LATE_FOLLOWUP_LIMIT = timedelta(days=2)
POLL_SECONDS = 30


def _fmt(dt):
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def sync_reminders(conn, event_id, settings, now=None):
    """Make the reminder rows of an active event match its offsets. Sent reminders are never touched."""
    now = now or datetime.now()
    event = conn.execute("SELECT * FROM events WHERE id = ?", (event_id,)).fetchone()
    if event is None or event["status"] != "active":
        return
    existing = {r["offset_token"]: r for r in
                conn.execute("SELECT * FROM reminders WHERE event_id = ?", (event_id,)).fetchall()}
    planned = planned_reminders(event, settings["all_day_send_time"])
    planned_tokens = set()

    for offset, send_at in planned:
        planned_tokens.add(offset.token)
        status, error = ("pending", None) if send_at > now else ("skipped", "Time had already passed when scheduled")
        row = existing.get(offset.token)
        if row is None:
            conn.execute(
                "INSERT INTO reminders (event_id, offset_label, offset_token, send_at, status, error) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (event_id, offset.label, offset.token, _fmt(send_at), status, error),
            )
        elif row["status"] not in ("sent", "sending"):
            conn.execute(
                "UPDATE reminders SET send_at = ?, status = ?, error = ?, attempts = 0, last_attempt_at = NULL "
                "WHERE id = ?",
                (_fmt(send_at), status, error, row["id"]),
            )

    for token, row in existing.items():
        if token not in planned_tokens and row["status"] not in ("sent", "sending"):
            conn.execute("DELETE FROM reminders WHERE id = ?", (row["id"],))


def _offset_from_token(token):
    offsets, _ = parse_offsets(token)
    return offsets[0] if offsets else Offset(0, "d")


def send_due(now=None):
    """Send every due reminder. Returns number sent."""
    now = now or datetime.now()
    sent = 0
    with db.get_db() as conn:
        settings = db.get_settings(conn)
        due = conn.execute(
            "SELECT r.* FROM reminders r JOIN events e ON e.id = r.event_id "
            "WHERE r.status = 'pending' AND e.status = 'active' AND r.send_at <= ? "
            "AND (r.last_attempt_at IS NULL OR r.last_attempt_at <= ?) ORDER BY r.send_at",
            (_fmt(now), _fmt(now - RETRY_AFTER)),
        ).fetchall()

    for reminder in due:
        with db.get_db() as conn:
            claimed = conn.execute(
                "UPDATE reminders SET status = 'sending', attempts = attempts + 1, last_attempt_at = ? "
                "WHERE id = ? AND status = 'pending'",
                (_fmt(now), reminder["id"]),
            ).rowcount
            event = conn.execute("SELECT * FROM events WHERE id = ?", (reminder["event_id"],)).fetchone()
        if not claimed or event is None:
            continue

        offset = _offset_from_token(reminder["offset_token"])
        send_at = datetime.strptime(reminder["send_at"], "%Y-%m-%d %H:%M:%S")
        missed = (
            (not offset.after and event_start(event, settings["all_day_send_time"]) <= now)
            or (offset.after and now - send_at > LATE_FOLLOWUP_LIMIT)
        )
        if missed:
            _finish(reminder["id"], "skipped", "Missed: the computer was off or asleep until it was too late")
            continue

        recipients, _ = parse_recipients(event["recipients"])
        subject, body = render_email(event, offset)
        try:
            method = mailer.send_email(settings, recipients, subject, body)
        except Exception as exc:
            attempts = reminder["attempts"] + 1
            status = "failed" if attempts >= MAX_ATTEMPTS else "pending"
            _finish(reminder["id"], status, str(exc))
            set_email_status(f"error: {exc}")
            logger.error("Reminder %s for %r failed (attempt %s): %s", reminder["id"], event["title"], attempts, exc)
            continue
        _finish(reminder["id"], "sent", None, sent_at=now)
        set_email_status("ok")
        sent += 1
        logger.info("Sent %s reminder for %r to %s via %s",
                    reminder["offset_label"], event["title"], ", ".join(recipients), method)
    return sent


def _finish(reminder_id, status, error, sent_at=None):
    with db.get_db() as conn:
        conn.execute(
            "UPDATE reminders SET status = ?, error = ?, sent_at = ? WHERE id = ?",
            (status, error, _fmt(sent_at) if sent_at else None, reminder_id),
        )


def set_email_status(status):
    with db.get_db() as conn:
        db.save_settings({"email_status": status}, conn)


def recover_interrupted():
    with db.get_db() as conn:
        conn.execute(
            "UPDATE reminders SET status = 'failed', "
            "error = 'Interrupted while sending. Retry if it did not arrive.' WHERE status = 'sending'"
        )


def _loop(stop_event):
    while not stop_event.is_set():
        try:
            send_due()
        except Exception:
            logger.exception("Error while sending due reminders")
        stop_event.wait(POLL_SECONDS)


def _catalog_loop(stop_event):
    import catalog

    last_attempt = None
    while not stop_event.is_set():
        try:
            due = catalog.sync_is_due(db.get_settings())
            if due and (last_attempt is None or datetime.now() - last_attempt >= timedelta(hours=1)):
                last_attempt = datetime.now()
                catalog.sync_catalog()
        except Exception:
            logger.exception("Error in catalogue sync loop")
        stop_event.wait(600)


def start_background_sender():
    recover_interrupted()
    stop_event = threading.Event()
    threading.Thread(target=_loop, args=(stop_event,), daemon=True, name="reminder-sender").start()
    threading.Thread(target=_catalog_loop, args=(stop_event,), daemon=True, name="catalog-loop").start()
    return stop_event
