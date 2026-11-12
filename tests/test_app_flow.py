"""End-to-end through the website: sign-in, per-user courses, admin settings (email sending is faked)."""

import importlib
import re
import sqlite3
from datetime import date, datetime, timedelta, timezone

import pytest

ADMIN = "chhavi.nayyar@ubc.ca"


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setenv("REMINDERS_DB", str(tmp_path / "test.db"))
    monkeypatch.setenv("APP_PASSWORD", "admin-secret")
    import db
    importlib.reload(db)
    import scheduler
    importlib.reload(scheduler)
    import catalog
    importlib.reload(catalog)
    import auth
    importlib.reload(auth)
    import app as app_module
    importlib.reload(app_module)

    sent = []
    fake_send = lambda settings, recipients, subject, body: sent.append((recipients, subject, body)) or "smtp"
    monkeypatch.setattr(scheduler.mailer, "send_email", fake_send)
    monkeypatch.setattr(app_module.mailer, "send_email", fake_send)
    monkeypatch.setattr(app_module.time, "sleep", lambda s: None)
    app_module.app.config.update(TESTING=True, CSRF_ENABLED=False)

    start = (date.today() + timedelta(days=30)).isoformat()
    end = (date.today() + timedelta(days=65)).isoformat()
    later = (date.today() + timedelta(days=150)).isoformat()
    with db.get_db() as conn:
        conn.execute("INSERT INTO catalog_courses (code, title, url, program) VALUES "
                     "('0115', 'Health Data and Visualization', 'u1', 'Health Data Analytics')")
        conn.execute("INSERT INTO catalog_courses (code, title, url, program) VALUES "
                     "('0576', 'NLP Information Session', 'u2', 'Natural Language Processing')")
        conn.execute("INSERT INTO catalog_sessions (section_code, course_code, start_date, end_date) "
                     "VALUES ('0115-010', '0115', ?, ?)", (start, end))
        conn.execute("INSERT INTO catalog_sessions (section_code, course_code, start_date) "
                     "VALUES ('0115-011', '0115', ?)", (later,))

    with app_module.app.test_client() as c:
        c.sent, c.db, c.scheduler, c.app = sent, db, scheduler, app_module
        c.start, c.end = start, end
        yield c


def sign_in_as(client, email):
    with client.session_transaction() as s:
        s.clear()
        s["user"] = email


def enable_email(client):
    with client.db.get_db() as conn:
        client.db.save_settings({"smtp_password": "pw", "email_status": "ok"}, conn)


# --- Sign-in --------------------------------------------------------------------------------------

def test_pages_need_sign_in(client):
    assert client.get("/health").status_code == 200
    assert "/login" in client.get("/").headers["Location"]
    assert client.post("/courses/add", data={"course": "0115"}).status_code == 302
    page = client.get("/login").get_data(as_text=True)
    assert "Sign in" in page and "How it works" in page and "mailto:chhavi.nayyar@ubc.ca" in page


def test_sign_in_survives_stale_session_cookie(client):
    enable_email(client)
    client.app.app.config["CSRF_ENABLED"] = True
    client.set_cookie("session", "eyJ1c2VyIjpudWxsfQ.stale.cookie")
    page = client.get("/login").get_data(as_text=True)
    token = re.search(r'name="csrf_token" value="([^"]+)"', page).group(1)
    page = client.post("/login/send", data={"csrf_token": token, "email": "pat@ubc.ca", "remember": "on"},
                       follow_redirects=True).get_data(as_text=True)
    assert "We sent a 6-digit code to" in page and len(client.sent) == 1


def test_email_code_sign_in(client):
    page = client.get("/login").get_data(as_text=True)
    assert "isn't switched on yet" in page and "Administrator sign-in" in page
    enable_email(client)
    assert "Email me a sign-in code" in client.get("/login").get_data(as_text=True)

    page = client.post("/login/send", data={"email": "someone@gmail.com"}, follow_redirects=True).get_data(as_text=True)
    assert "Please use your UBC email address" in page and not client.sent

    page = client.post("/login/send", data={"email": " Pat.Lee@UBC.ca "}, follow_redirects=True).get_data(as_text=True)
    assert "We sent a 6-digit code to" in page and "<strong>pat.lee@ubc.ca</strong>" in page
    assert "Resend code" in page and "Change email" in page
    [(recipients, subject, body)] = client.sent
    assert recipients == ["pat.lee@ubc.ca"]
    code = re.search(r"code is: (\d{6})", body).group(1)
    assert subject.endswith(code) and "/login/link?token=" in body

    wrong = "000000" if code != "000000" else "111111"
    assert "didn&#39;t work" in client.post("/login/verify", data={"code": wrong}, follow_redirects=True).get_data(as_text=True)
    page = client.post("/login/verify", data={"code": code}, follow_redirects=True).get_data(as_text=True)
    assert "pat.lee@ubc.ca" in page and "Your courses" in page
    assert "Settings" not in page and "People using it" not in page

    client.post("/logout")
    assert "/login" in client.get("/").headers["Location"]
    client.post("/login/send", data={"email": "pat.lee@ubc.ca"})
    client.post("/login/restart")
    assert "/login" in client.post("/login/verify", data={"code": code}).headers["Location"]


def test_sign_in_link_works_once(client):
    enable_email(client)
    client.post("/login/send", data={"email": "sam@student.ubc.ca"})
    link = re.search(r"(/login/link\?token=\S+)", client.sent[-1][2]).group(1)
    client.post("/logout")
    assert client.get(link).headers["Location"].endswith("/")
    assert "sam@student.ubc.ca" in client.get("/").get_data(as_text=True)
    client.post("/logout")
    assert "expired" in client.get(link, follow_redirects=True).get_data(as_text=True)


def test_code_attempts_and_requests_are_limited(client):
    enable_email(client)
    for _ in range(5):
        client.post("/login/send", data={"email": "pat@ubc.ca"})
    page = client.post("/login/send", data={"email": "pat@ubc.ca"}, follow_redirects=True).get_data(as_text=True)
    assert "Too many sign-in emails" in page
    code = re.search(r"code is: (\d{6})", client.sent[-1][2]).group(1)
    for _ in range(5):
        client.post("/login/verify", data={"code": "999999" if code != "999999" else "888888"})
    client.post("/login/verify", data={"code": code})
    assert "/login" in client.get("/").headers["Location"]


def session_cookie(resp):
    return next(c for c in resp.headers.getlist("Set-Cookie") if c.startswith("cr_session="))


def sign_in_with_code(client, email, remember):
    data = {"email": email}
    if remember:
        data["remember"] = "on"
    client.post("/login/send", data=data)
    code = re.search(r"code is: (\d{6})", client.sent[-1][2]).group(1)
    return client.post("/login/verify", data={"code": code})


def test_remember_me_keeps_you_signed_in_for_90_days(client):
    enable_email(client)
    resp = sign_in_with_code(client, "pat@ubc.ca", remember=True)
    cookie = session_cookie(resp)
    expires = re.search(r"Expires=([^;]+)", cookie).group(1)
    days = (datetime.strptime(expires, "%a, %d %b %Y %H:%M:%S GMT") - datetime.now(timezone.utc).replace(tzinfo=None)).days
    assert 88 <= days <= 90
    assert "Expires=" in session_cookie(client.get("/"))


def test_without_remember_me_session_ends_with_browser(client):
    enable_email(client)
    page = client.get("/login").get_data(as_text=True)
    assert '<input type="checkbox" name="remember" checked>' in page
    client.post("/login/send", data={"email": "pat@ubc.ca"})
    assert "signed out when you close your browser" in client.get("/login").get_data(as_text=True)
    code = re.search(r"code is: (\d{6})", client.sent[-1][2]).group(1)
    resp = client.post("/login/verify", data={"code": code})
    assert "Expires=" not in session_cookie(resp)
    assert "pat@ubc.ca" in client.get("/").get_data(as_text=True)

    resp = client.post("/login/admin", data={"password": "admin-secret"})
    assert "Expires=" not in session_cookie(resp)
    resp = client.post("/login/admin", data={"password": "admin-secret", "remember": "on"})
    assert "Expires=" in session_cookie(resp)


def test_last_email_is_remembered_and_can_be_forgotten(client):
    enable_email(client)
    resp = sign_in_with_code(client, "pat@ubc.ca", remember=True)
    last = next(c for c in resp.headers.getlist("Set-Cookie") if c.startswith("last_email="))
    assert "pat@ubc.ca" in last and "HttpOnly" in last
    client.post("/logout")
    page = client.get("/login").get_data(as_text=True)
    assert 'value="pat@ubc.ca"' in page and "Not you?" in page and "Use a different email" in page

    client.get("/login/forget")
    page = client.get("/login").get_data(as_text=True)
    assert 'value="pat@ubc.ca"' not in page and "Not you?" not in page


def test_resend_code_and_change_email(client):
    enable_email(client)
    client.post("/login/send", data={"email": "pat@ubc.ca", "remember": "on"})
    first = re.search(r"code is: (\d{6})", client.sent[-1][2]).group(1)
    page = client.post("/login/resend", follow_redirects=True).get_data(as_text=True)
    assert "We sent a new code to pat@ubc.ca" in page and len(client.sent) == 2
    second = re.search(r"code is: (\d{6})", client.sent[-1][2]).group(1)
    if first != second:
        assert "/login" in client.post("/login/verify", data={"code": first}).headers["Location"]
    for _ in range(3):
        client.post("/login/resend")
    page = client.post("/login/resend", follow_redirects=True).get_data(as_text=True)
    assert "Too many sign-in emails" in page
    client.post("/login/restart")
    assert "Email me a sign-in code" in client.get("/login").get_data(as_text=True)


def test_admin_password_sign_in(client):
    assert "Wrong administrator" in client.post("/login/admin", data={"password": "nope"},
                                                follow_redirects=True).get_data(as_text=True)
    page = client.post("/login/admin", data={"password": "admin-secret"}, follow_redirects=True).get_data(as_text=True)
    assert ADMIN in page and "Set up sending" in page and "Settings" in page


def test_forms_need_csrf_token(client):
    client.app.app.config["CSRF_ENABLED"] = True
    sign_in_as(client, "pat@ubc.ca")
    client.post("/courses/add", data={"course": ["0115"], "recipients": "pat@ubc.ca", "offsets": "1d"})
    with client.db.get_db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM followed_courses").fetchone()[0] == 0
    page = client.get("/").get_data(as_text=True)
    token = re.search(r'name="csrf_token" value="([^"]+)"', page).group(1)
    client.post("/courses/add", data={"course": ["0115"], "recipients": "pat@ubc.ca", "offsets": "1d",
                                      "csrf_token": token})
    with client.db.get_db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM followed_courses").fetchone()[0] == 1


def test_can_be_embedded_in_ubc_sites(client):
    csp = client.get("/login").headers["Content-Security-Policy"]
    assert "frame-ancestors" in csp and "https://*.ubc.ca" in csp and "https://sites.olt.ubc.ca" in csp


# --- Courses --------------------------------------------------------------------------------------

def test_add_course_turns_reminders_on(client):
    sign_in_as(client, ADMIN)
    page = client.get("/").get_data(as_text=True)
    assert f'name="recipients" value="{ADMIN}"' in page
    assert page.index('id="add-button"') < page.index('id="course-search"')
    assert "Getting started" in page and '<section class="intro">' in page
    resp = client.post("/courses/add", data={"course": ["0115"], "recipients": "me@ubc.ca, team@ubc.ca",
                                             "offsets": "14d, 2d"}, follow_redirects=True)
    page = resp.get_data(as_text=True)
    assert "Reminders are on for Health Data and Visualization (2 upcoming sessions)" in page
    assert 'name="program"' not in page and "Save changes" not in page

    with client.db.get_db() as conn:
        events = conn.execute("SELECT * FROM events ORDER BY start_date").fetchall()
        reminders = conn.execute("SELECT * FROM reminders WHERE event_id = ? ORDER BY send_at",
                                 (events[0]["id"],)).fetchall()
    assert [(e["status"], e["owner"]) for e in events] == [("active", ADMIN)] * 2

    two_days_before = datetime.strptime(reminders[1]["send_at"], "%Y-%m-%d %H:%M:%S")
    assert client.scheduler.send_due(now=two_days_before) == 2
    recipients, subject, body = client.sent[-1]
    assert recipients == ["me@ubc.ca", "team@ubc.ca"]
    assert subject == "Course Reminder: Health Data and Visualization"
    assert body == (
        "Dear Team,\n\n"
        "This is a reminder that Health Data and Visualization is starting in 2 days.\n\n"
        "Course Details:\n"
        "- Course: Health Data and Visualization\n"
        "- Program: Health Data\n"
        f"- Start Date: {client.start}\n"
        f"- End Date: {client.end}\n"
    )


def test_users_only_see_and_manage_their_own_courses(client):
    sign_in_as(client, "pat@ubc.ca")
    client.post("/courses/add", data={"course": ["0115"], "recipients": "pat-team@ubc.ca", "offsets": "7d"})
    sign_in_as(client, "sam@ubc.ca")
    client.post("/courses/add", data={"course": ["0115", "0576"], "recipients": "sam-team@ubc.ca", "offsets": "1d"})

    with client.db.get_db() as conn:
        follows = {(r["owner"], r["course_code"]): r["recipients"] for r in conn.execute("SELECT * FROM followed_courses")}
        pat_event = conn.execute("SELECT id FROM events WHERE owner = 'pat@ubc.ca'").fetchone()["id"]
        counts = dict(conn.execute("SELECT owner, COUNT(*) FROM events GROUP BY owner").fetchall())
    assert follows == {("pat@ubc.ca", "0115"): "pat-team@ubc.ca", ("sam@ubc.ca", "0115"): "sam-team@ubc.ca",
                       ("sam@ubc.ca", "0576"): "sam-team@ubc.ca"}
    assert counts == {"pat@ubc.ca": 2, "sam@ubc.ca": 2}

    page = client.get("/").get_data(as_text=True)
    assert 'name="owner" value="sam@ubc.ca"' in page and 'name="owner" value="pat@ubc.ca"' not in page
    assert client.post("/courses/0115/edit", data={"owner": "pat@ubc.ca", "recipients": "x@ubc.ca",
                                                   "offsets": "1d"}).status_code == 403
    assert client.post("/courses/0115/remove", data={"owner": "pat@ubc.ca"}).status_code == 403
    assert client.post(f"/sessions/{pat_event}/remove").status_code == 404
    assert client.post("/settings", data={"action": "save", "smtp_password": "x"}).status_code == 403
    assert client.post("/courses/refresh").status_code == 403

    client.post("/courses/0115/remove")
    with client.db.get_db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM events WHERE owner = 'pat@ubc.ca'").fetchone()[0] == 2
        assert conn.execute("SELECT COUNT(*) FROM events WHERE owner = 'sam@ubc.ca'").fetchone()[0] == 0

    sign_in_as(client, ADMIN)
    page = client.get("/").get_data(as_text=True)
    assert "Everyone else's courses" in page and "Owner: pat@ubc.ca" in page and "Owner: sam@ubc.ca" in page
    client.post("/courses/0115/edit", data={"owner": "pat@ubc.ca", "recipients": "boss@ubc.ca", "offsets": "3d",
                                            "program": "HD"})
    with client.db.get_db() as conn:
        rows = conn.execute("SELECT recipients, offsets, program FROM events WHERE owner = 'pat@ubc.ca'").fetchall()
    assert {tuple(r) for r in rows} == {("boss@ubc.ca", "3d", "HD")}


def test_skip_session_and_bad_input(client):
    sign_in_as(client, "pat@ubc.ca")
    page = client.post("/courses/add", data={"course": ["0115"], "recipients": "nope", "offsets": "soon"},
                       follow_redirects=True).get_data(as_text=True)
    assert "aren&#39;t valid email addresses" in page
    client.post("/courses/add", data={"course": ["0115"], "recipients": "pat@ubc.ca", "offsets": "1d"})
    with client.db.get_db() as conn:
        first = conn.execute("SELECT id FROM events ORDER BY start_date").fetchone()["id"]
    client.post(f"/sessions/{first}/remove")
    with client.db.get_db() as conn:
        assert [tuple(r) for r in conn.execute("SELECT owner, section_code FROM dismissed_sessions")] == [
            ("pat@ubc.ca", "0115-010")]


def test_admin_saves_email_password_and_gets_test_email(client):
    sign_in_as(client, ADMIN)
    page = client.post("/settings", data={"action": "save", "smtp_password": "pw"},
                       follow_redirects=True).get_data(as_text=True)
    assert "Email works" in page and "Set up sending" not in page
    assert client.sent[-1][0] == [ADMIN]


def test_dashboard_summary_counts(client):
    sign_in_as(client, "pat@ubc.ca")
    client.post("/courses/add", data={"course": ["0115"], "recipients": "pat@ubc.ca", "offsets": "7d, 1d"})
    page = client.get("/").get_data(as_text=True)
    assert "<span>1 course</span>" in page and "<span>4 emails scheduled</span>" in page
    assert "<span>0 sent</span>" in page and "Next reminder:" in page and "Health Data and Visualization" in page


# --- Migration ------------------------------------------------------------------------------------

def test_single_user_database_is_migrated_to_admin(tmp_path, monkeypatch):
    path = tmp_path / "old.db"
    old = sqlite3.connect(path)
    old.executescript("""
        CREATE TABLE events (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, start_date TEXT,
            start_time TEXT NOT NULL DEFAULT '', end_date TEXT NOT NULL DEFAULT '', end_time TEXT NOT NULL DEFAULT '',
            location TEXT NOT NULL DEFAULT '', notes TEXT NOT NULL DEFAULT '', recipients TEXT NOT NULL DEFAULT '',
            offsets TEXT NOT NULL DEFAULT '', status TEXT NOT NULL DEFAULT 'draft',
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')), source_key TEXT, course_code TEXT,
            program TEXT NOT NULL DEFAULT '');
        CREATE TABLE followed_courses (course_code TEXT PRIMARY KEY, recipients TEXT NOT NULL, offsets TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')), program TEXT NOT NULL DEFAULT '');
        CREATE TABLE dismissed_sessions (section_code TEXT PRIMARY KEY);
        INSERT INTO followed_courses (course_code, recipients, offsets, program)
            VALUES ('0116', 'chhavi.nayyar@ubc.ca', '14d, 7d, 2d, 1d', 'Health Data'),
                   ('0576', 'chhavi.nayyar@ubc.ca', '14d, 7d, 2d, 1d', '');
        INSERT INTO events (title, start_date, status, source_key, course_code)
            VALUES ('Health Data ML', '2027-11-17', 'active', '0116-009', '0116');
        INSERT INTO dismissed_sessions VALUES ('0116-008');
    """)
    old.close()

    monkeypatch.setenv("REMINDERS_DB", str(path))
    import db
    importlib.reload(db)
    db.init_db()
    db.init_db()
    with db.get_db() as conn:
        follows = [tuple(r) for r in conn.execute(
            "SELECT owner, course_code, recipients, program FROM followed_courses ORDER BY course_code")]
        events = [tuple(r) for r in conn.execute("SELECT owner, source_key FROM events")]
        dismissed = [tuple(r) for r in conn.execute("SELECT owner, section_code FROM dismissed_sessions")]
        conn.execute("INSERT INTO followed_courses (owner, course_code, recipients, offsets) "
                     "VALUES ('pat@ubc.ca', '0116', 'pat@ubc.ca', '1d')")
    assert follows == [(ADMIN, "0116", "chhavi.nayyar@ubc.ca", "Health Data"), (ADMIN, "0576", "chhavi.nayyar@ubc.ca", "")]
    assert events == [(ADMIN, "0116-009")]
    assert dismissed == [(ADMIN, "0116-008")]
