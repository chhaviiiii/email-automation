#!/usr/bin/env python3
"""Course reminders: pick UBC Extended Learning courses and reminder emails go out automatically.

People sign in with their UBC email (a code is emailed to them). Everyone manages their own courses;
reminder emails are all sent from the shared sender address. Admins also see everyone's courses and
the settings.
"""

import hmac
import logging
import os
import secrets
import time
from datetime import date, datetime, timedelta

from flask import Flask, abort, flash, redirect, render_template, request, session, url_for
from werkzeug.middleware.proxy_fix import ProxyFix

import auth
import catalog
import db
import mailer
import scheduler
from reminder_logic import parse_offsets, parse_recipients

APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
ON_FLY = bool(os.environ.get("FLY_APP_NAME"))
FRAME_ANCESTORS = os.environ.get("FRAME_ANCESTORS", "'self' https://*.ubc.ca https://sites.olt.ubc.ca")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler(os.path.join(os.path.dirname(db.DB_PATH), "reminders.log")),
              logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

db.init_db()
with db.get_db() as _conn:
    catalog.fill_missing_programs(_conn)

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
app.secret_key = db.get_settings()["secret_key"]
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    # Inside a WordPress iframe the site is third-party, so on HTTPS the cookie must be SameSite=None.
    SESSION_COOKIE_SAMESITE="None" if ON_FLY else "Lax",
    SESSION_COOKIE_SECURE=ON_FLY,
    SESSION_COOKIE_PARTITIONED=ON_FLY,
    PERMANENT_SESSION_LIFETIME=timedelta(days=90),
    CSRF_ENABLED=True,
)

LAST_EMAIL_COOKIE = "last_email"
PUBLIC_ENDPOINTS = {"login", "login_send", "login_verify", "login_link", "login_admin", "login_restart",
                    "login_resend", "login_forget", "health", "static"}


# --- Security plumbing --------------------------------------------------------------------------

def csrf_token():
    if "csrf" not in session:
        session["csrf"] = secrets.token_urlsafe(32)
    return session["csrf"]


def current_user():
    return session.get("user")


def is_admin():
    return auth.is_admin(current_user())


@app.before_request
def protect():
    if request.method == "POST" and app.config["CSRF_ENABLED"] and request.endpoint != "health":
        sent = request.form.get("csrf_token", "")
        if not session.get("csrf") or not hmac.compare_digest(sent, session["csrf"]):
            flash("Your browser didn't keep you signed in on this page. Try again, or open Course Reminders "
                  "in its own tab.", "error")
            return redirect(url_for("index" if current_user() else "login"))
    if request.endpoint in PUBLIC_ENDPOINTS or current_user():
        return None
    return redirect(url_for("login"))


@app.after_request
def security_headers(resp):
    resp.headers["Content-Security-Policy"] = f"frame-ancestors {FRAME_ANCESTORS}"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    return resp


# --- Sign in ------------------------------------------------------------------------------------

def email_ready(settings):
    return bool(mailer.get_password(settings)) and not settings["email_status"].startswith("error")


def wants_remember():
    return request.form.get("remember") == "on"


def sign_in(email, remember):
    """Start a session. Remembered sessions last 90 days (renewed on each visit); others end with the browser."""
    email = auth.normalize(email)
    session.clear()
    session.permanent = bool(remember)
    session["user"] = email
    logger.info("Signed in: %s (remember=%s)", email, bool(remember))
    resp = redirect(url_for("index"))
    resp.set_cookie(LAST_EMAIL_COOKIE, email, max_age=int(timedelta(days=365).total_seconds()), httponly=True,
                    secure=ON_FLY, samesite="None" if ON_FLY else "Lax", partitioned=ON_FLY)
    return resp


def remembered_email():
    email = auth.normalize(request.cookies.get(LAST_EMAIL_COOKIE))
    return email if auth.email_allowed(email) else ""


@app.get("/login")
def login():
    if current_user():
        return redirect(url_for("index"))
    settings = db.get_settings()
    return render_template(
        "login.html", pending=session.get("pending_email"), pending_remember=session.get("pending_remember", True),
        email_ready=email_ready(settings), admin_login=bool(APP_PASSWORD), domains=auth.ALLOWED_DOMAINS,
        last_email=remembered_email(),
    )


def send_code(email, remember):
    """Email a sign-in code to email and move the sign-in page to the code step."""
    with db.get_db() as conn:
        settings = db.get_settings(conn)
        if not mailer.get_password(settings):
            flash("Email sign-in isn't switched on yet. Please try again later.", "error")
            return redirect(url_for("login"))
        try:
            code, token = auth.create_login(conn, email, app.secret_key, remember)
        except auth.LoginError as exc:
            flash(str(exc), "error")
            return redirect(url_for("login"))
    link = (os.environ.get("APP_URL") or request.url_root.rstrip("/")) + url_for("login_link", token=token)
    try:
        auth.send_login_email(settings, email, code, link)
    except Exception as exc:
        logger.error("Couldn't send sign-in email to %s: %s", email, exc)
        flash("We couldn't send the sign-in email right now. Please try again in a few minutes.", "error")
        return redirect(url_for("login"))
    session["pending_email"] = email
    session["pending_remember"] = bool(remember)
    return redirect(url_for("login"))


@app.post("/login/send")
def login_send():
    email = auth.normalize(request.form.get("email"))
    if not auth.email_allowed(email):
        flash(f"Please use your UBC email address (ending in {', '.join('@' + d for d in auth.ALLOWED_DOMAINS)}).",
              "error")
        return redirect(url_for("login"))
    return send_code(email, wants_remember())


@app.post("/login/resend")
def login_resend():
    email = session.get("pending_email")
    if not email:
        return redirect(url_for("login"))
    resp = send_code(email, session.get("pending_remember", True))
    if session.get("pending_email"):
        flash(f"We sent a new code to {email}.", "success")
    return resp


@app.post("/login/verify")
def login_verify():
    email = session.get("pending_email")
    if not email:
        return redirect(url_for("login"))
    with db.get_db() as conn:
        ok = auth.verify_code(conn, email, request.form.get("code", ""), app.secret_key)
    if not ok:
        time.sleep(1)
        flash("That code didn't work. Check the latest email, or request a new code.", "error")
        return redirect(url_for("login"))
    return sign_in(email, session.get("pending_remember", True))


@app.get("/login/link")
def login_link():
    with db.get_db() as conn:
        email, remember = auth.verify_token(conn, request.args.get("token", ""), app.secret_key)
    if not email:
        flash("That sign-in link has expired or was already used. Request a new one.", "error")
        return redirect(url_for("login"))
    return sign_in(email, remember)


@app.post("/login/restart")
def login_restart():
    session.pop("pending_email", None)
    session.pop("pending_remember", None)
    return redirect(url_for("login"))


@app.get("/login/forget")
def login_forget():
    resp = redirect(url_for("login"))
    resp.delete_cookie(LAST_EMAIL_COOKIE, secure=ON_FLY, samesite="None" if ON_FLY else "Lax", partitioned=ON_FLY)
    return resp


@app.post("/login/admin")
def login_admin():
    if APP_PASSWORD and hmac.compare_digest(request.form.get("password", "").encode(), APP_PASSWORD.encode()):
        return sign_in(db.DEFAULT_OWNER, wants_remember())
    time.sleep(1)
    flash("Wrong administrator password.", "error")
    return redirect(url_for("login"))


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.get("/health")
def health():
    return {"status": "ok"}


# --- Display helpers ----------------------------------------------------------------------------

app.jinja_env.globals["csrf_token"] = csrf_token


@app.context_processor
def inject_globals():
    return {"user": current_user(), "admin": is_admin()}


@app.template_filter("dt")
def format_datetime(value):
    if not value:
        return ""
    if isinstance(value, str):
        value = datetime.strptime(value[:16], "%Y-%m-%d %H:%M")
    hour = value.hour % 12 or 12
    return f"{value:%a, %b} {value.day}, {value.year}, {hour}:{value.minute:02d} {'AM' if value.hour < 12 else 'PM'}"


def _as_datetime(value):
    return datetime.strptime(value[:16], "%Y-%m-%d %H:%M") if isinstance(value, str) else value


@app.template_filter("day")
def format_day(value):
    value = _as_datetime(value)
    return f"{value:%a, %b} {value.day}" if value else ""


@app.template_filter("time12")
def format_time12(value):
    value = _as_datetime(value)
    return f"{value.hour % 12 or 12}:{value.minute:02d} {'AM' if value.hour < 12 else 'PM'}" if value else ""


@app.template_filter("daterange")
def format_daterange(row):
    start = date.fromisoformat(row["start_date"])
    end = date.fromisoformat(row["end_date"]) if row["end_date"] else None
    if end and end != start:
        first = f"{start:%b} {start.day}" + (f", {start.year}" if start.year != end.year else "")
        return f"{first} – {end:%b} {end.day}, {end.year}"
    return f"{start:%b} {start.day}, {start.year}"


def email_problem(settings):
    if not mailer.get_password(settings):
        return "missing"
    if settings["email_status"].startswith("error"):
        return settings["email_status"][len("error: "):]
    return None


def check_offsets_and_recipients(form):
    recipients, bad_recipients = parse_recipients(form.get("recipients", ""))
    offsets, bad_offsets = parse_offsets(form.get("offsets", ""))
    errors = []
    if bad_recipients:
        errors.append("These aren't valid email addresses: " + ", ".join(bad_recipients))
    elif not recipients:
        errors.append("Add at least one email address to send reminders to.")
    if bad_offsets or not offsets:
        errors.append("Reminder times should look like: 14d, 7d, 2d, 1d (d = days, w = weeks, h = hours).")
    return ", ".join(recipients), ", ".join(o.token for o in offsets), errors


def course_cards(conn, follows, today):
    cards = []
    for f in follows:
        sessions = []
        for e in conn.execute(
            "SELECT * FROM events WHERE owner = ? AND course_code = ? AND status = 'active' "
            "AND COALESCE(NULLIF(end_date, ''), start_date) >= ? ORDER BY start_date",
            (f["owner"], f["course_code"], today),
        ).fetchall():
            reminders = conn.execute("SELECT * FROM reminders WHERE event_id = ? ORDER BY send_at",
                                     (e["id"],)).fetchall()
            sessions.append({
                "event": e,
                "next": next((r for r in reminders if r["status"] == "pending"), None),
                "sent": sum(r["status"] == "sent" for r in reminders),
                "failed": [r for r in reminders if r["status"] == "failed"],
            })
        cards.append({"course": f, "sessions": sessions})
    return cards


def summary(conn, owner=None):
    where, params = ("WHERE e.owner = ?", (owner,)) if owner else ("", ())
    follows = conn.execute(
        "SELECT COUNT(*) FROM followed_courses" + (" WHERE owner = ?" if owner else ""), params).fetchone()[0]
    row = conn.execute(
        "SELECT SUM(r.status = 'pending' AND e.status = 'active') AS scheduled, SUM(r.status = 'sent') AS sent, "
        "MIN(CASE WHEN r.status = 'pending' AND e.status = 'active' THEN r.send_at END) AS next_send "
        f"FROM reminders r JOIN events e ON e.id = r.event_id {where}", params,
    ).fetchone()
    return {"courses": follows, "scheduled": row["scheduled"] or 0, "sent": row["sent"] or 0,
            "next": row["next_send"]}


def can_manage(owner):
    return owner == current_user() or is_admin()


# --- Main page ----------------------------------------------------------------------------------

@app.route("/")
def index():
    today = date.today().isoformat()
    user = current_user()
    with db.get_db() as conn:
        settings = db.get_settings(conn)
        follow_query = ("SELECT f.*, c.title, c.url FROM followed_courses f "
                        "LEFT JOIN catalog_courses c ON c.code = f.course_code ")
        my_courses = course_cards(conn, conn.execute(
            follow_query + "WHERE f.owner = ? ORDER BY c.title", (user,)).fetchall(), today)
        others = []
        if is_admin():
            others = course_cards(conn, conn.execute(
                follow_query + "WHERE f.owner != ? ORDER BY f.owner, c.title", (user,)).fetchall(), today)

        followed_codes = {c["course"]["course_code"] for c in my_courses}
        upcoming = {}
        for s in conn.execute(catalog.upcoming_sessions_query(), (today,)).fetchall():
            upcoming.setdefault(s["course_code"], []).append(s)
        available = [
            {"course": c, "sessions": upcoming.get(c["code"], [])}
            for c in conn.execute("SELECT * FROM catalog_courses ORDER BY title").fetchall()
            if c["code"] not in followed_codes
        ]
        stats = summary(conn, None if is_admin() else user)
        people = conn.execute("SELECT COUNT(DISTINCT owner) FROM followed_courses").fetchone()[0]

    return render_template(
        "index.html", settings=settings, my_courses=my_courses, others=others, available=available,
        stats=stats, people=people, email_problem=email_problem(settings), syncing=catalog.sync_running(),
    )


@app.post("/courses/add")
def add_courses():
    user = current_user()
    codes = list(dict.fromkeys(request.form.getlist("course")))
    recipients, offsets, errors = check_offsets_and_recipients(request.form)
    if not codes:
        errors.append("Tick at least one course to add.")
    if errors:
        for e in errors:
            flash(e, "error")
        return redirect(url_for("index") + "#add")

    with db.get_db() as conn:
        settings = db.get_settings(conn)
        added_titles, sessions = [], 0
        for code in codes:
            course = conn.execute("SELECT * FROM catalog_courses WHERE code = ?", (code,)).fetchone()
            if course is None:
                continue
            conn.execute(
                "INSERT OR REPLACE INTO followed_courses (owner, course_code, recipients, offsets, program) "
                "VALUES (?, ?, ?, ?, ?)",
                (user, code, recipients, offsets, catalog.default_program(conn, code)),
            )
            conn.execute("DELETE FROM dismissed_sessions WHERE owner = ? AND section_code LIKE ?", (user, code + "-%"))
            sessions += catalog.add_course_sessions(conn, user, code, settings, include_dismissed=True)
            added_titles.append(course["title"])

    if added_titles:
        flash(f"Reminders are on for {', '.join(added_titles)}"
              + (f" ({sessions} upcoming session{'s' if sessions != 1 else ''})." if sessions else
                 ". No dates are posted yet; reminders start automatically when they are."), "success")
    return redirect(url_for("index"))


def form_owner():
    owner = auth.normalize(request.form.get("owner")) or current_user()
    if not can_manage(owner):
        abort(403)
    return owner


@app.post("/courses/<code>/edit")
def edit_course(code):
    owner = form_owner()
    recipients, offsets, errors = check_offsets_and_recipients(request.form)
    if errors:
        for e in errors:
            flash(e, "error")
        return redirect(url_for("index"))
    program = request.form.get("program", "").strip()
    with db.get_db() as conn:
        settings = db.get_settings(conn)
        conn.execute("UPDATE followed_courses SET recipients = ?, offsets = ?, program = ? "
                     "WHERE owner = ? AND course_code = ?", (recipients, offsets, program, owner, code))
        events = conn.execute("SELECT id FROM events WHERE owner = ? AND course_code = ? AND status = 'active'",
                              (owner, code)).fetchall()
        for e in events:
            conn.execute("UPDATE events SET recipients = ?, offsets = ?, program = ? WHERE id = ?",
                         (recipients, offsets, program, e["id"]))
            scheduler.sync_reminders(conn, e["id"], settings)
    flash("Saved.", "success")
    return redirect(url_for("index"))


@app.post("/courses/<code>/remove")
def remove_course(code):
    owner = form_owner()
    with db.get_db() as conn:
        conn.execute("DELETE FROM followed_courses WHERE owner = ? AND course_code = ?", (owner, code))
        conn.execute("DELETE FROM events WHERE owner = ? AND course_code = ?", (owner, code))
    flash("Reminders turned off for that course.", "success")
    return redirect(url_for("index"))


def owned_event(conn, event_id):
    event = conn.execute("SELECT * FROM events WHERE id = ?", (event_id,)).fetchone()
    if event is None or not can_manage(event["owner"]):
        abort(404)
    return event


@app.post("/sessions/<int:event_id>/remove")
def remove_session(event_id):
    with db.get_db() as conn:
        event = owned_event(conn, event_id)
        if event["source_key"]:
            conn.execute("INSERT OR IGNORE INTO dismissed_sessions (owner, section_code) VALUES (?, ?)",
                         (event["owner"], event["source_key"]))
        conn.execute("DELETE FROM events WHERE id = ?", (event_id,))
    flash("Reminders turned off for that session.", "success")
    return redirect(url_for("index"))


@app.post("/reminders/<int:reminder_id>/retry")
def retry_reminder(reminder_id):
    with db.get_db() as conn:
        reminder = conn.execute("SELECT event_id FROM reminders WHERE id = ?", (reminder_id,)).fetchone()
        if reminder is None:
            abort(404)
        owned_event(conn, reminder["event_id"])
        conn.execute(
            "UPDATE reminders SET status = 'pending', attempts = 0, last_attempt_at = NULL, error = NULL, "
            "send_at = MIN(send_at, ?) WHERE id = ? AND status = 'failed'",
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), reminder_id),
        )
    flash("Trying again within a minute.", "success")
    return redirect(url_for("index"))


# --- Admin --------------------------------------------------------------------------------------

def require_admin():
    if not is_admin():
        abort(403)


@app.post("/courses/refresh")
def refresh_courses():
    require_admin()
    catalog.sync_in_background()
    flash("Updating the course list from Extended Learning. This takes a minute or two.", "success")
    return redirect(url_for("index") + "#add")


@app.post("/settings")
def save_settings():
    require_admin()
    action = request.form.get("action", "save")
    password = request.form.get("smtp_password", "")
    with db.get_db() as conn:
        settings = db.get_settings(conn)
        if action == "save":
            values = {}
            if "offsets" in request.form:
                offsets, bad = parse_offsets(request.form.get("offsets", ""))
                send_time = request.form.get("all_day_send_time", "09:00")
                errors = []
                if bad or not offsets:
                    errors.append("Reminder times should look like: 14d, 7d, 2d, 1d.")
                try:
                    datetime.strptime(send_time, "%H:%M")
                except ValueError:
                    errors.append("Email time should look like 09:00.")
                if errors:
                    for e in errors:
                        flash(e, "error")
                    return redirect(url_for("index") + "#settings")
                values.update(default_offsets=", ".join(o.token for o in offsets), all_day_send_time=send_time)
            if password:
                values["smtp_password"] = password
            db.save_settings(values, conn)
            if values.get("all_day_send_time", settings["all_day_send_time"]) != settings["all_day_send_time"]:
                new_settings = db.get_settings(conn)
                for e in conn.execute("SELECT id FROM events WHERE status = 'active'").fetchall():
                    scheduler.sync_reminders(conn, e["id"], new_settings)
            settings = db.get_settings(conn)
        if action == "test" or (action == "save" and password):
            try:
                mailer.send_email(settings, [current_user()], "Course reminders are working",
                                  "This is a test from the Course Reminders website. Reminder and sign-in emails "
                                  f"will be sent from {settings['sender_email']}.\n")
                db.save_settings({"email_status": "ok"}, conn)
                flash(f"Email works. A test email was sent to {current_user()}.", "success")
            except Exception as exc:
                db.save_settings({"email_status": f"error: {exc}"}, conn)
                flash(f"Couldn't send email: {exc}", "error")
        elif action == "save":
            flash("Settings saved.", "success")
    return redirect(url_for("index"))


if __name__ == "__main__":
    scheduler.start_background_sender()
    port = int(os.environ.get("PORT", "5055"))
    logger.info("Reminder website running at http://127.0.0.1:%s", port)
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)
