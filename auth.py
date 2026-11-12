"""Email sign-in: a 6-digit code (and a one-click link) sent from the shared UBC sender address."""

import hashlib
import hmac
import os
import re
import secrets
from datetime import datetime, timedelta

import db
import mailer

ALLOWED_DOMAINS = [d.strip().lower().lstrip("@") for d in
                   os.environ.get("ALLOWED_EMAIL_DOMAINS", "ubc.ca").split(",") if d.strip()]
CODE_TTL = timedelta(minutes=15)
MAX_CODES_PER_HOUR = 5
MAX_ATTEMPTS = 5
EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+(?:\.[\w-]+)+$")


class LoginError(Exception):
    pass


def normalize(email):
    return (email or "").strip().lower()


def is_admin(email):
    return normalize(email) in db.ADMIN_EMAILS


def email_allowed(email):
    email = normalize(email)
    if not EMAIL_RE.match(email):
        return False
    if is_admin(email):
        return True
    domain = email.rsplit("@", 1)[1]
    return any(domain == d or domain.endswith("." + d) for d in ALLOWED_DOMAINS)


def _hash(value, secret):
    return hmac.new(secret.encode(), value.encode(), hashlib.sha256).hexdigest()


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def create_login(conn, email, secret, remember=True):
    """Store a new code + link token for email. Returns (code, token)."""
    since = (datetime.now() - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
    recent = conn.execute("SELECT COUNT(*) FROM login_codes WHERE email = ? AND created_at >= ?",
                          (email, since)).fetchone()[0]
    if recent >= MAX_CODES_PER_HOUR:
        raise LoginError("Too many sign-in emails were requested. Please wait an hour and try again.")
    code = f"{secrets.randbelow(1_000_000):06d}"
    token = secrets.token_urlsafe(32)
    conn.execute(
        "INSERT INTO login_codes (email, code_hash, token_hash, expires_at, created_at, remember) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (email, _hash(code, secret), _hash(token, secret),
         (datetime.now() + CODE_TTL).strftime("%Y-%m-%d %H:%M:%S"), _now(), int(bool(remember))),
    )
    return code, token


def verify_code(conn, email, code, secret):
    """True if code matches the newest unused, unexpired code for email."""
    row = conn.execute(
        "SELECT * FROM login_codes WHERE email = ? AND used_at IS NULL AND expires_at >= ? "
        "ORDER BY id DESC LIMIT 1", (email, _now())
    ).fetchone()
    if row is None or row["attempts"] >= MAX_ATTEMPTS:
        return False
    if hmac.compare_digest(row["code_hash"], _hash(re.sub(r"\D", "", code or ""), secret)):
        conn.execute("UPDATE login_codes SET used_at = ? WHERE id = ?", (_now(), row["id"]))
        return True
    conn.execute("UPDATE login_codes SET attempts = attempts + 1 WHERE id = ?", (row["id"],))
    return False


def verify_token(conn, token, secret):
    """(email, remember) for a valid one-click sign-in link, or (None, False)."""
    row = conn.execute(
        "SELECT * FROM login_codes WHERE token_hash = ? AND used_at IS NULL AND expires_at >= ?",
        (_hash(token or "", secret), _now()),
    ).fetchone()
    if row is None:
        return None, False
    conn.execute("UPDATE login_codes SET used_at = ? WHERE id = ?", (_now(), row["id"]))
    return row["email"], bool(row["remember"])


def send_login_email(settings, email, code, link):
    body = (
        f"Your Course Reminders sign-in code is: {code}\n\n"
        f"Enter it on the sign-in page, or open this link to sign in:\n{link}\n\n"
        f"The code and link expire in {int(CODE_TTL.total_seconds() // 60)} minutes. "
        "If you didn't ask to sign in, you can ignore this email.\n"
    )
    mailer.send_email(settings, [email], f"Course Reminders sign-in code: {code}", body)
