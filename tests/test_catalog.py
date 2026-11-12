"""Catalogue parsing (markup mirrors extendedlearning.ubc.ca) and sync behaviour with a fake website."""

import importlib
from datetime import date

import pytest

TODAY = date(2026, 10, 29)

LISTING = """
<table><tbody>
<tr class="relative grid js-reveal reveal">
  <td><button><span class="field field--name-title">Health Data and Visualization</span></button></td>
  <td><div class="field__item">Online</div></td>
  <td> $1,025 </td><td> Mar 02, 2027 </td><td> Open </td>
  <td><a href="/courses/health-data-visualization/0115" rel="bookmark">Details</a></td>
  <td class="js-reveal__target" hidden> Insights into real-time clinical data. <p><strong>Other start dates:</strong> TBA</p></td>
</tr>
<tr class="relative grid js-reveal reveal">
  <td><button><span class="field field--name-title">Health Data Analysis and Machine Learning</span></button></td>
  <td><div class="field__item">Online</div></td>
  <td> $1,025 </td><td> Nov 17, 2026 </td><td> Open </td>
  <td><a href="/courses/health-data-analysis-machine-learning/0116" rel="bookmark">Details</a></td>
  <td class="js-reveal__target" hidden> Machine learning for health data.</td>
</tr>
</tbody></table>
<nav><a href="?page=0">1</a><a href="?page=1">2</a></nav>
"""

LISTING_PAGE_2 = """
<table><tbody>
<tr class="js-reveal">
  <td><span class="field--name-title">Fluency Under Pressure</span></td>
  <td>Online</td><td>$500</td><td>Oct 28, 2026</td><td>Open</td>
  <td><a href="/courses/fluency-under-pressure/0614" rel="bookmark">Details</a></td>
</tr>
</tbody></table>
"""


def course_page(*sessions):
    rows = "".join(f"""
<tr class="relative grid js-reveal reveal">
  <td><button><span class="md--hidden">Starting </span> {dates} </button></td>
  <td> {days} </td><td> {time} </td><td> Online </td><td> $1,025 </td>
  <td><button class="js-button--cart" data-toast-title="{code}">Add to cart</button></td>
  <td class="js-reveal__target" hidden><h3>Code</h3><p>{code}</p></td>
</tr>""" for code, dates, days, time in sessions)
    return f"""<section><h2>Available sessions</h2><table>
<caption class="sr-only">Available course sessions</caption>
<thead><tr><th>Dates</th></tr></thead><tbody>{rows}</tbody></table></section>
<h2>Related courses</h2><table><tbody><tr><td>Nov 17, 2026</td></tr></tbody></table>"""


def test_parse_listing():
    import catalog
    courses = catalog.parse_listing(LISTING)
    assert [(c["code"], c["title"], c["format"]) for c in courses] == [
        ("0115", "Health Data and Visualization", "Online"),
        ("0116", "Health Data Analysis and Machine Learning", "Online"),
    ]
    assert courses[0]["url"] == "https://extendedlearning.ubc.ca/courses/health-data-visualization/0115"
    assert courses[0]["summary"] == "Insights into real-time clinical data."
    assert catalog.last_listing_page(LISTING) == 1


def test_parse_course_page():
    import catalog
    html = course_page(("0115-010", "Mar 02, 2027 to Apr 05, 2027", "-", "-"),
                       ("0614-001", "Oct 28, 2026 to Dec 09, 2026", "Wed", "5:30 pm"))
    sessions = catalog.parse_course_page(html, "0115", TODAY)
    assert [(s["section_code"], s["start_date"], s["end_date"], s["start_time"], s["days"]) for s in sessions] == [
        ("0115-010", "2027-03-02", "2027-04-05", "", ""),
        ("0614-001", "2026-10-28", "2026-12-09", "17:30", "Wed"),
    ]
    assert catalog.parse_course_page("<p>No sessions</p>", "0115", TODAY) == []


class FakeSite:
    def __init__(self):
        self.pages = {
            "page=0": LISTING,
            "page=1": LISTING_PAGE_2,
            "/0115": course_page(("0115-010", "Mar 02, 2027 to Apr 05, 2027", "-", "-")),
            "/0116": course_page(("0116-009", "Nov 17, 2026 to Dec 21, 2026", "-", "-")),
            "/0614": course_page(),
        }

    def __call__(self, url):
        for key, html in self.pages.items():
            if url.endswith(key):
                return html
        raise AssertionError(url)


@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.setenv("REMINDERS_DB", str(tmp_path / "test.db"))
    import db
    importlib.reload(db)
    import scheduler
    importlib.reload(scheduler)
    import catalog
    importlib.reload(catalog)
    db.init_db()
    return db, catalog


def follow(conn, code, recipients="a@ubc.ca", offsets="7d", program="", owner="a@ubc.ca"):
    conn.execute("INSERT INTO followed_courses (owner, course_code, recipients, offsets, program) "
                 "VALUES (?, ?, ?, ?, ?)", (owner, code, recipients, offsets, program))


def test_sync_adds_sessions_of_followed_courses_and_updates_dates(env):
    db, catalog = env
    site = FakeSite()
    site.pages["/0116"] = course_page(("0116-009", "Nov 17, 2027 to Dec 21, 2027", "-", "-"))
    catalog.sync_catalog(fetch=site, today=TODAY)
    assert db.get_settings()["catalog_sync_status"] == "ok: 3 courses"

    with db.get_db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM catalog_sessions").fetchone()[0] == 2
        follow(conn, "0116", "b@ubc.ca", "14d, 1d", owner="b@ubc.ca")
        assert catalog.add_course_sessions(conn, "b@ubc.ca", "0116", db.get_settings(conn), TODAY) == 1
        assert catalog.add_course_sessions(conn, "b@ubc.ca", "0116", db.get_settings(conn), TODAY) == 0
        follow(conn, "0115")

    site.pages["/0115"] = course_page(("0115-010", "Mar 02, 2027 to Apr 05, 2027", "-", "-"),
                                      ("0115-011", "Jun 01, 2027 to Jul 05, 2027", "-", "-"))
    site.pages["/0116"] = course_page(("0116-009", "Nov 24, 2027 to Dec 28, 2027", "-", "-"))
    catalog.sync_catalog(fetch=site, today=TODAY)

    with db.get_db() as conn:
        events = {e["source_key"]: e for e in conn.execute("SELECT * FROM events")}
        moved = conn.execute("SELECT send_at FROM reminders WHERE event_id = ? ORDER BY send_at",
                             (events["0116-009"]["id"],)).fetchall()
    assert set(events) == {"0115-010", "0115-011", "0116-009"}
    assert {e["status"] for e in events.values()} == {"active"}
    assert events["0115-011"]["recipients"] == "a@ubc.ca"
    assert events["0115-011"]["title"] == "Health Data and Visualization"
    assert events["0115-011"]["program"] == "Health Data"
    assert events["0116-009"]["start_date"] == "2027-11-24"
    assert [r["send_at"] for r in moved] == ["2027-11-10 09:00:00", "2027-11-23 09:00:00"]


def test_dismissed_sessions_are_not_re_added(env):
    db, catalog = env
    catalog.sync_catalog(fetch=FakeSite(), today=TODAY)
    with db.get_db() as conn:
        follow(conn, "0115")
        conn.execute("INSERT INTO dismissed_sessions (owner, section_code) VALUES ('a@ubc.ca', '0115-010')")
        assert catalog.apply_catalog_changes(conn, db.get_settings(conn), TODAY) == (0, 0)
        follow(conn, "0115", owner="c@ubc.ca")
        assert catalog.apply_catalog_changes(conn, db.get_settings(conn), TODAY) == (0, 1)
        owners = [r["owner"] for r in conn.execute("SELECT owner FROM events WHERE source_key = '0115-010'")]
    assert owners == ["c@ubc.ca"]


def test_parse_program():
    import catalog
    html = """<article><p>This course is part of the <a href="/programs-credentials/nlp-microcertificate">UBC
    Micro-certificate in Natural Language</a> Processing program.</p></article>
    <a href="/programs-credentials/other">Other program</a>"""
    assert catalog.parse_program_url(html) == "https://extendedlearning.ubc.ca/programs-credentials/nlp-microcertificate"
    assert catalog.parse_program_url("<article>No program</article>") == ""
    assert catalog.parse_program_name("<h1> Natural Language Processing to Improve Patient Care </h1>") == (
        "Natural Language Processing to Improve Patient Care")


def test_sync_reads_program_names(env):
    db, catalog = env
    site = FakeSite()
    site.pages["/0614"] = '<article><a href="/programs-credentials/business-english">x</a></article>'
    site.pages["/business-english"] = "<h1>Business English for International Professionals</h1>"
    catalog.sync_catalog(fetch=site, today=TODAY)
    with db.get_db() as conn:
        assert conn.execute("SELECT program FROM catalog_courses WHERE code = '0614'").fetchone()[0] == (
            "Business English for International Professionals")
