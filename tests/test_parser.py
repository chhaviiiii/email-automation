from datetime import date

import pandas as pd

from schedule_parser import parse_free_text, parse_rows, parse_text

TODAY = date(2026, 10, 29)  # a Thursday


def events(text):
    return [e.to_dict() for e in parse_free_text(text, TODAY).events]


def test_month_day_with_time_and_location():
    [e] = events("Nov 5, 2pm - Team meeting @ ICICS 104")
    assert e["title"] == "Team meeting"
    assert e["start_date"] == "2026-11-05"
    assert e["start_time"] == "14:00"
    assert e["location"] == "ICICS 104"


def test_date_range_rolls_into_next_year():
    [e] = events("DASH 101: Jan 6 – Feb 14")
    assert e["title"] == "DASH 101"
    assert (e["start_date"], e["end_date"]) == ("2027-01-06", "2027-02-14")
    assert e["start_time"] == ""


def test_weekday_prefix_and_time_range():
    [e] = events("Thursday Nov 12 10-11:30am Advising session")
    assert e["title"] == "Advising session"
    assert (e["start_date"], e["start_time"], e["end_time"]) == ("2026-11-12", "10:00", "11:30")


def test_full_date_with_time_range_and_room():
    [e] = events("Midterm: Feb 20, 2027, 6:30–8:30 PM, Room SWNG 121")
    assert (e["title"], e["location"]) == ("Midterm", "Room SWNG 121")
    assert (e["start_date"], e["start_time"], e["end_time"]) == ("2027-02-20", "18:30", "20:30")


def test_iso_date_and_course_code_is_kept():
    [e] = events("2026-12-01 DATA 501 grades due")
    assert e["title"] == "DATA 501 grades due"
    assert e["start_date"] == "2026-12-01"


def test_date_header_with_timed_items():
    result = events("Monday, Nov 16\n- 9am Standup\n- 3pm Project review")
    assert [(e["title"], e["start_date"], e["start_time"]) for e in result] == [
        ("Standup", "2026-11-16", "09:00"),
        ("Project review", "2026-11-16", "15:00"),
    ]


def test_title_line_followed_by_date_line_and_field_lines():
    [e] = events("Thesis defense\nDec 3 at 1:30pm\nLocation: Room 204\nTo: prof@ubc.ca")
    assert e["title"] == "Thesis defense"
    assert (e["start_date"], e["start_time"]) == ("2026-12-03", "13:30")
    assert e["location"] == "Room 204"
    assert e["recipients"] == "prof@ubc.ca"


def test_relative_dates_and_emails():
    [a, b] = events("tomorrow 10am dentist\nnext Friday lunch with sam@ubc.ca")
    assert (a["title"], a["start_date"]) == ("dentist", "2026-10-30")
    assert (b["title"], b["start_date"], b["recipients"]) == ("lunch with", "2026-10-30", "sam@ubc.ca")


def test_numeric_and_day_month_dates():
    [a, b] = events("11/20 Workshop\n5 December 2026 Holiday party")
    assert (a["title"], a["start_date"]) == ("Workshop", "2026-11-20")
    assert (b["title"], b["start_date"]) == ("Holiday party", "2026-12-05")


def test_undated_lines_are_reported():
    result = parse_free_text("Fall 2026 schedule\nNov 5 Kickoff", TODAY)
    assert [e.title for e in result.events] == ["Kickoff"]
    assert result.skipped == ["Fall 2026 schedule"]


def test_pasted_excel_table():
    text = "Course Name\tProgram Name\tStart Date\tEnd Date\nNLP 1\tNLP\t2026-11-10\t2026-12-15\nKCDS 2\tKCDS\t1/12/2027\t"
    result = parse_text(text, TODAY)
    assert [(e.title, e.start_date, e.end_date) for e in result.events] == [
        ("NLP 1", "2026-11-10", "2026-12-15"),
        ("KCDS 2", "2027-01-12", ""),
    ]
    assert result.events[0].notes == "Program Name: NLP"


def test_pasted_table_with_dashed_dates():
    text = (
        "Program Name\tD1 Term Code\tCourse Number & Section\tCourse Name\tStart Date\tEnd Date\n"
        "Health Data\t2026-27 Winter Term 1\t0114-009 \tIntro to Big Data\t5-Jan-2027\t8-Feb-2027\n"
        "Health Data\t2026-27 Winter Term 1\t0115-009 \tHealth Data and Visualization\t2-Mar-27\t5-Apr-27\n"
    )
    result = parse_text(text, TODAY)
    assert [(e.title, e.start_date, e.end_date) for e in result.events] == [
        ("Intro to Big Data", "2027-01-05", "2027-02-08"),
        ("Health Data and Visualization", "2027-03-02", "2027-04-05"),
    ]
    assert result.skipped == []


def test_dashed_dates_in_free_text():
    [a, b] = events("Kickoff 14-Jan-2027\nReview Jan-20-2027 3pm")
    assert (a["title"], a["start_date"]) == ("Kickoff", "2027-01-14")
    assert (b["title"], b["start_date"], b["start_time"]) == ("Review", "2027-01-20", "15:00")


def test_spreadsheet_rows_with_timestamps():
    df = pd.DataFrame({
        "Event": ["Kickoff", None],
        "Date": [pd.Timestamp("2026-11-03 09:30"), None],
        "Room": ["ICICS 104", None],
    })
    [e] = parse_rows(df, TODAY).events
    assert (e.title, e.start_date, e.start_time, e.location) == ("Kickoff", "2026-11-03", "09:30", "ICICS 104")
