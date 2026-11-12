"""Course catalogue from UBC Extended Learning (extendedlearning.ubc.ca).

A daily sync reads the course listing and each course page's "Available sessions" table into
catalog_courses / catalog_sessions. Courses you add get reminders for every upcoming session; after each
sync, new sessions of those courses get reminders too, and changed dates are updated.
"""

import logging
import re
import threading
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta

from bs4 import BeautifulSoup

import db
from schedule_parser import parse_line

logger = logging.getLogger(__name__)

BASE_URL = "https://extendedlearning.ubc.ca"
LISTING_URL = BASE_URL + "/courses-programs/our-courses?page={page}"
USER_AGENT = "UBC-course-reminders/1.0 (chhavi.nayyar@ubc.ca)"
SYNC_EVERY = timedelta(hours=24)
WORKERS = 4

# Program names used in the team's original course schedule spreadsheet.
PROGRAM_NAMES = {
    "0100": "Software Development",
    "0101": "Data Science", "0102": "Data Science", "0103": "Data Science", "0104": "Data Science",
    "0114": "Health Data", "0115": "Health Data", "0116": "Health Data",
}

_sync_lock = threading.Lock()


def fetch_html(url, retries=2):
    last_exc = None
    for _ in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=30) as resp:
                return resp.read().decode("utf-8", errors="replace")
        except Exception as exc:
            last_exc = exc
    raise last_exc


def _text(node):
    return re.sub(r"\s+", " ", node.get_text(" ", strip=True)).strip() if node else ""


def parse_listing(html):
    """Courses on one listing page: dicts with code, title, url, format, summary."""
    soup = BeautifulSoup(html, "html.parser")
    courses = []
    for row in soup.select("tr.js-reveal"):
        link = row.find("a", rel="bookmark")
        title = row.select_one(".field--name-title")
        if not link or not title or "/courses/" not in link.get("href", ""):
            continue
        cells = row.find_all("td", recursive=False)
        summary = re.sub(r"\s*Other start dates:.*$", "", _text(row.select_one("td.js-reveal__target")))
        href = link["href"]
        courses.append({
            "code": href.rstrip("/").rsplit("/", 1)[-1],
            "title": _text(title),
            "url": href if href.startswith("http") else BASE_URL + href,
            "format": _text(cells[1]) if len(cells) > 1 else "",
            "summary": summary,
            "program": "",
        })
    return courses


def last_listing_page(html):
    pages = [int(n) for n in re.findall(r'href="\?page=(\d+)"', html)]
    return max(pages) if pages else 0


def parse_program_url(html):
    """URL of the program a course belongs to (linked from "This course is part of the ... program")."""
    article = BeautifulSoup(html, "html.parser").find("article")
    link = article.find("a", href=re.compile(r"/programs-credentials/")) if article else None
    if not link:
        return ""
    href = link["href"].strip()
    return href if href.startswith("http") else BASE_URL + href


def parse_program_name(html):
    return _text(BeautifulSoup(html, "html.parser").find("h1"))


def parse_course_page(html, course_code, today=None):
    """Dated sessions from a course page's "Available course sessions" table."""
    today = today or date.today()
    soup = BeautifulSoup(html, "html.parser")
    caption = soup.find("caption", string=re.compile("Available course sessions", re.I))
    if not caption:
        return []
    sessions = []
    for row in caption.find_parent("table").select("tbody > tr"):
        cells = row.find_all("td", recursive=False)
        if len(cells) < 4:
            continue
        dates_text = re.sub(r"^Starting\s+", "", _text(cells[0]), flags=re.I)
        cart = row.find(attrs={"data-toast-title": True})
        code = cart["data-toast-title"].strip() if cart else ""
        if not code:
            code_heading = row.find(["h3", "h4"], string=re.compile(r"^\s*Code\s*$"))
            code = _text(code_heading.find_next("p")) if code_heading else ""
        if not code:
            continue
        parsed = parse_line(dates_text, today)
        if not parsed["dates"]:
            continue
        times = parse_line(_text(cells[2]), today)["times"]
        sessions.append({
            "section_code": code,
            "course_code": course_code,
            "start_date": parsed["dates"][0].isoformat(),
            "end_date": parsed["dates"][-1].isoformat() if len(parsed["dates"]) > 1 else "",
            "start_time": times[0].strftime("%H:%M") if times else "",
            "end_time": times[1].strftime("%H:%M") if len(times) > 1 else "",
            "days": _text(cells[1]).strip(" -"),
            "format": _text(cells[3]),
            "dates_text": dates_text,
        })
    return sessions


def default_program(conn, course_code):
    if course_code in PROGRAM_NAMES:
        return PROGRAM_NAMES[course_code]
    row = conn.execute("SELECT program FROM catalog_courses WHERE code = ?", (course_code,)).fetchone()
    return row["program"] if row else ""


def upcoming_sessions_query(where="", order="s.start_date"):
    return (
        "SELECT s.*, c.title, c.url FROM catalog_sessions s JOIN catalog_courses c ON c.code = s.course_code "
        "WHERE COALESCE(NULLIF(s.end_date, ''), s.start_date) >= ? " + where + f" ORDER BY {order}"
    )


def add_session_event(conn, section_code, course, settings):
    """Turn on reminders for one catalogue session using a followed course's settings. Returns event id or None."""
    import scheduler

    if conn.execute("SELECT 1 FROM events WHERE source_key = ? AND owner = ? AND status IN ('active', 'paused')",
                    (section_code, course["owner"])).fetchone():
        return None
    session = conn.execute(
        "SELECT s.*, c.title, c.format AS course_format FROM catalog_sessions s "
        "JOIN catalog_courses c ON c.code = s.course_code WHERE s.section_code = ?", (section_code,)
    ).fetchone()
    if session is None:
        return None
    values = {
        "title": session["title"],
        "start_date": session["start_date"],
        "start_time": session["start_time"],
        "end_date": session["end_date"],
        "end_time": session["end_time"],
        "location": session["format"] or session["course_format"],
        "notes": f"Section: {section_code}",
        "recipients": course["recipients"],
        "offsets": course["offsets"],
        "program": course["program"],
        "status": "active",
        "source_key": section_code,
        "course_code": session["course_code"],
        "owner": course["owner"],
    }
    cur = conn.execute(
        f"INSERT INTO events ({', '.join(values)}) VALUES ({', '.join('?' for _ in values)})",
        list(values.values()),
    )
    scheduler.sync_reminders(conn, cur.lastrowid, settings)
    return cur.lastrowid


def add_course_sessions(conn, owner, course_code, settings, today=None, include_dismissed=False):
    """Turn on reminders for every upcoming session of a course someone follows. Returns number added."""
    course = conn.execute("SELECT * FROM followed_courses WHERE owner = ? AND course_code = ?",
                          (owner, course_code)).fetchone()
    if course is None:
        return 0
    where, params = "AND s.course_code = ?", [course_code]
    if not include_dismissed:
        where += " AND s.section_code NOT IN (SELECT section_code FROM dismissed_sessions WHERE owner = ?)"
        params.append(owner)
    today = (today or date.today()).isoformat()
    rows = conn.execute(upcoming_sessions_query(where), (today, *params)).fetchall()
    return sum(1 for s in rows if add_session_event(conn, s["section_code"], course, settings))


def fill_missing_programs(conn):
    for row in conn.execute("SELECT owner, course_code FROM followed_courses WHERE program = ''").fetchall():
        owner, code = row["owner"], row["course_code"]
        program = default_program(conn, code)
        if program:
            conn.execute("UPDATE followed_courses SET program = ? WHERE owner = ? AND course_code = ?",
                         (program, owner, code))
            conn.execute("UPDATE events SET program = ? WHERE owner = ? AND course_code = ? AND program = ''",
                         (program, owner, code))


def apply_catalog_changes(conn, settings, today=None):
    """Update events whose session dates changed and add new sessions of followed courses.

    Returns (changed, added) counts.
    """
    import scheduler

    fill_missing_programs(conn)
    changed = 0
    linked = conn.execute(
        "SELECT e.*, s.start_date AS s_start, s.end_date AS s_end, s.start_time AS s_stime, s.end_time AS s_etime "
        "FROM events e JOIN catalog_sessions s ON s.section_code = e.source_key "
        "WHERE e.status IN ('active', 'paused')"
    ).fetchall()
    for e in linked:
        new = (e["s_start"], e["s_end"], e["s_stime"], e["s_etime"])
        if new != (e["start_date"], e["end_date"], e["start_time"], e["end_time"]):
            conn.execute(
                "UPDATE events SET start_date = ?, end_date = ?, start_time = ?, end_time = ? WHERE id = ?",
                (*new, e["id"]),
            )
            scheduler.sync_reminders(conn, e["id"], settings)
            changed += 1
            logger.info("%s (%s) dates changed to %s – %s", e["title"], e["source_key"], new[0], new[1])

    added = sum(add_course_sessions(conn, f["owner"], f["course_code"], settings, today)
                for f in conn.execute("SELECT owner, course_code FROM followed_courses").fetchall())
    return changed, added


def _set_status(status, synced=False):
    values = {"catalog_sync_status": status}
    if synced:
        values["catalog_synced_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with db.get_db() as conn:
        db.save_settings(values, conn)


def sync_catalog(fetch=fetch_html, today=None):
    """Download the catalogue and apply changes. Returns False if a sync was already running."""
    if not _sync_lock.acquire(blocking=False):
        return False
    try:
        _set_status("running")
        first = fetch(LISTING_URL.format(page=0))
        courses = {c["code"]: c for c in parse_listing(first)}
        for page in range(1, last_listing_page(first) + 1):
            for c in parse_listing(fetch(LISTING_URL.format(page=page))):
                courses.setdefault(c["code"], c)
        if not courses:
            raise RuntimeError("No courses found on the Extended Learning listing pages")

        def course_sessions(course):
            try:
                html = fetch(course["url"])
                course["program_url"] = parse_program_url(html)
                return parse_course_page(html, course["code"], today)
            except Exception as exc:
                logger.warning("Couldn't read %s: %s", course["url"], exc)
                return None

        def program_name(url):
            try:
                return url, parse_program_name(fetch(url))
            except Exception as exc:
                logger.warning("Couldn't read %s: %s", url, exc)
                return url, ""

        with ThreadPoolExecutor(WORKERS) as pool:
            results = list(pool.map(course_sessions, courses.values()))
            program_urls = {c["program_url"] for c in courses.values() if c.get("program_url")}
            programs = dict(pool.map(program_name, program_urls))
        for c in courses.values():
            c["program"] = programs.get(c.get("program_url"), "")
        failures = sum(r is None for r in results)
        if failures > len(courses) / 2:
            raise RuntimeError(f"Couldn't read {failures} of {len(courses)} course pages")

        with db.get_db() as conn:
            conn.execute("DELETE FROM catalog_courses")
            conn.execute("DELETE FROM catalog_sessions")
            for c in courses.values():
                conn.execute(
                    "INSERT INTO catalog_courses (code, title, url, format, summary, program) VALUES (?, ?, ?, ?, ?, ?)",
                    (c["code"], c["title"], c["url"], c["format"], c["summary"], c["program"]),
                )
            for sessions in results:
                for s in sessions or []:
                    conn.execute(
                        "INSERT OR REPLACE INTO catalog_sessions (section_code, course_code, start_date, end_date, "
                        "start_time, end_time, days, format, dates_text) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (s["section_code"], s["course_code"], s["start_date"], s["end_date"], s["start_time"],
                         s["end_time"], s["days"], s["format"], s["dates_text"]),
                    )
            changed, added = apply_catalog_changes(conn, db.get_settings(conn), today)

        note = f" ({failures} course pages couldn't be read)" if failures else ""
        _set_status(f"ok: {len(courses)} courses{note}", synced=True)
        logger.info("Catalogue synced: %s courses, %s changed sessions, %s new sessions", len(courses), changed, added)
        return True
    except Exception as exc:
        logger.exception("Catalogue sync failed")
        _set_status(f"error: {exc}")
        return True
    finally:
        _sync_lock.release()


def sync_running():
    return _sync_lock.locked()


def sync_in_background():
    if sync_running():
        return False
    threading.Thread(target=sync_catalog, daemon=True, name="catalog-sync").start()
    return True


def sync_is_due(settings):
    if not settings.get("catalog_synced_at"):
        return True
    last = datetime.strptime(settings["catalog_synced_at"], "%Y-%m-%d %H:%M:%S")
    return datetime.now() - last >= SYNC_EVERY
