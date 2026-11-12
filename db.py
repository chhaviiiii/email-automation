"""SQLite storage for events, reminders and settings."""

import os
import secrets
import sqlite3
from contextlib import contextmanager

DB_PATH = os.environ.get(
    "REMINDERS_DB", os.path.join(os.path.dirname(os.path.abspath(__file__)), "reminders.db")
)

ADMIN_EMAILS = [e.strip().lower() for e in os.environ.get("ADMIN_EMAILS", "chhavi.nayyar@ubc.ca").split(",")
                if e.strip()]
DEFAULT_OWNER = ADMIN_EMAILS[0]

DEFAULT_SETTINGS = {
    "sender_email": "chhavi.nayyar@ubc.ca",
    "smtp_server": "smtp.mail.ubc.ca",
    "smtp_port": "587",
    "default_offsets": "14d, 7d, 2d, 1d",
    "all_day_send_time": "09:00",
    "catalog_synced_at": "",
    "catalog_sync_status": "",
    "smtp_password": "",
    "email_status": "",
}

SCHEMA = """
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    start_date TEXT,
    start_time TEXT NOT NULL DEFAULT '',
    end_date TEXT NOT NULL DEFAULT '',
    end_time TEXT NOT NULL DEFAULT '',
    location TEXT NOT NULL DEFAULT '',
    notes TEXT NOT NULL DEFAULT '',
    recipients TEXT NOT NULL DEFAULT '',
    offsets TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT 'draft',
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE TABLE IF NOT EXISTS reminders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_id INTEGER NOT NULL REFERENCES events(id) ON DELETE CASCADE,
    offset_label TEXT NOT NULL,
    offset_token TEXT NOT NULL,
    send_at TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'pending',
    attempts INTEGER NOT NULL DEFAULT 0,
    last_attempt_at TEXT,
    sent_at TEXT,
    error TEXT
);

CREATE INDEX IF NOT EXISTS idx_reminders_due ON reminders(status, send_at);

CREATE TABLE IF NOT EXISTS catalog_courses (
    code TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    url TEXT NOT NULL,
    format TEXT NOT NULL DEFAULT '',
    summary TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS catalog_sessions (
    section_code TEXT PRIMARY KEY,
    course_code TEXT NOT NULL,
    start_date TEXT NOT NULL,
    end_date TEXT NOT NULL DEFAULT '',
    start_time TEXT NOT NULL DEFAULT '',
    end_time TEXT NOT NULL DEFAULT '',
    days TEXT NOT NULL DEFAULT '',
    format TEXT NOT NULL DEFAULT '',
    dates_text TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS followed_courses (
    owner TEXT NOT NULL,
    course_code TEXT NOT NULL,
    recipients TEXT NOT NULL,
    offsets TEXT NOT NULL,
    program TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    PRIMARY KEY (owner, course_code)
);

CREATE TABLE IF NOT EXISTS dismissed_sessions (
    owner TEXT NOT NULL,
    section_code TEXT NOT NULL,
    PRIMARY KEY (owner, section_code)
);

CREATE TABLE IF NOT EXISTS login_codes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT NOT NULL,
    code_hash TEXT NOT NULL,
    token_hash TEXT NOT NULL,
    expires_at TEXT NOT NULL,
    attempts INTEGER NOT NULL DEFAULT 0,
    used_at TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);
"""


def connect():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def get_db():
    conn = connect()
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with get_db() as conn:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.executescript(SCHEMA)
        _add_column(conn, "events", "source_key", "TEXT")
        _add_column(conn, "events", "course_code", "TEXT")
        _add_column(conn, "events", "program", "TEXT NOT NULL DEFAULT ''")
        _add_column(conn, "events", "owner", "TEXT")
        _add_column(conn, "followed_courses", "program", "TEXT NOT NULL DEFAULT ''")
        _add_column(conn, "catalog_courses", "program", "TEXT NOT NULL DEFAULT ''")
        _add_column(conn, "login_codes", "remember", "INTEGER NOT NULL DEFAULT 1")
        conn.execute(
            "UPDATE events SET course_code = substr(source_key, 1, instr(source_key, '-') - 1) "
            "WHERE course_code IS NULL AND source_key LIKE '%-%'"
        )
        conn.execute("UPDATE events SET owner = ? WHERE owner IS NULL", (DEFAULT_OWNER,))
        _migrate_to_per_user_tables(conn)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_events_owner ON events(owner, course_code)")
        for key, value in DEFAULT_SETTINGS.items():
            conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (key, value))
        conn.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES ('secret_key', ?)",
            (secrets.token_hex(32),),
        )


def _columns(conn, table):
    return {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}


def _migrate_to_per_user_tables(conn):
    """Single-user databases had one follow / dismissal per course; they now belong to the admin."""
    if "owner" not in _columns(conn, "followed_courses"):
        conn.execute("ALTER TABLE followed_courses RENAME TO followed_courses_old")
        conn.execute(
            "CREATE TABLE followed_courses (owner TEXT NOT NULL, course_code TEXT NOT NULL, recipients TEXT NOT NULL, "
            "offsets TEXT NOT NULL, program TEXT NOT NULL DEFAULT '', "
            "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')), PRIMARY KEY (owner, course_code))"
        )
        conn.execute(
            "INSERT INTO followed_courses (owner, course_code, recipients, offsets, program, created_at) "
            "SELECT ?, course_code, recipients, offsets, program, created_at FROM followed_courses_old",
            (DEFAULT_OWNER,),
        )
        conn.execute("DROP TABLE followed_courses_old")
    if "owner" not in _columns(conn, "dismissed_sessions"):
        conn.execute("ALTER TABLE dismissed_sessions RENAME TO dismissed_sessions_old")
        conn.execute("CREATE TABLE dismissed_sessions (owner TEXT NOT NULL, section_code TEXT NOT NULL, "
                     "PRIMARY KEY (owner, section_code))")
        conn.execute("INSERT INTO dismissed_sessions (owner, section_code) "
                     "SELECT ?, section_code FROM dismissed_sessions_old", (DEFAULT_OWNER,))
        conn.execute("DROP TABLE dismissed_sessions_old")


def _add_column(conn, table, column, definition):
    if column not in {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def get_settings(conn=None):
    if conn is None:
        with get_db() as c:
            return get_settings(c)
    rows = conn.execute("SELECT key, value FROM settings").fetchall()
    settings = dict(DEFAULT_SETTINGS)
    settings.update({r["key"]: r["value"] for r in rows})
    return settings


def save_settings(values, conn):
    for key, value in values.items():
        conn.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, value),
        )
