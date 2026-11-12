from datetime import datetime

from reminder_logic import parse_offsets, parse_recipients, planned_reminders, render_email


def make_event(**overrides):
    event = {"id": 1, "title": "Team meeting", "start_date": "2026-11-05", "start_time": "", "end_date": "",
             "end_time": "", "location": "", "notes": "", "recipients": "a@ubc.ca", "offsets": "14d, 7d, 2d, 1d",
             "program": ""}
    event.update(overrides)
    return event


def test_parse_offsets():
    offsets, bad = parse_offsets("14d, 1 week, 2h, 30 minutes, 14d after, soon")
    assert [o.token for o in offsets] == ["14d", "1w", "2h", "30m", "14d after"]
    assert bad == ["soon"]


def test_parse_recipients():
    valid, invalid = parse_recipients("a@ubc.ca; B@ubc.ca, a@ubc.ca nope")
    assert valid == ["a@ubc.ca", "B@ubc.ca"]
    assert invalid == ["nope"]


def test_all_day_event_reminders_go_out_at_send_time():
    plans = planned_reminders(make_event(offsets="7d, 1d"), "09:00")
    assert [(o.label, t) for o, t in plans] == [
        ("7 days before", datetime(2026, 10, 29, 9, 0)),
        ("1 day before", datetime(2026, 11, 4, 9, 0)),
    ]


def test_timed_event_and_follow_up():
    event = make_event(start_time="14:00", end_date="2026-12-10", offsets="2h, 14d after")
    plans = planned_reminders(event, "09:00")
    assert [t for _, t in plans] == [datetime(2026, 11, 5, 12, 0), datetime(2026, 12, 24, 14, 0)]


def test_render_email_matches_team_format():
    event = make_event(title="Health Data and Visualization", program="Health Data",
                       start_date="2025-10-07", end_date="2025-11-10")
    subject, body = render_email(event, parse_offsets("2d")[0][0])
    assert subject == "Course Reminder: Health Data and Visualization"
    assert body == (
        "Dear Team,\n\n"
        "This is a reminder that Health Data and Visualization is starting in 2 days.\n\n"
        "Course Details:\n"
        "- Course: Health Data and Visualization\n"
        "- Program: Health Data\n"
        "- Start Date: 2025-10-07\n"
        "- End Date: 2025-11-10\n"
    )


def test_render_email_wording():
    event = make_event(start_time="17:00")
    phrases = {t: render_email(event, parse_offsets(t)[0][0])[1].splitlines()[2]
               for t in ["14d", "1d", "0d", "14d after"]}
    assert phrases["14d"] == "This is a reminder that Team meeting is starting in 2 weeks."
    assert phrases["1d"] == "This is a reminder that Team meeting is starting tomorrow."
    assert phrases["0d"] == "This is a reminder that Team meeting is starting today."
    assert phrases["14d after"] == "This is a follow-up reminder for Team meeting, which ended 14 days ago."
    body = render_email(event, parse_offsets("1d")[0][0])[1]
    assert "- Program: N/A" in body and "- End Date: N/A" in body and "- Time: 5:00 PM" in body
