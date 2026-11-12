"""Reminder offsets, send-time calculation and email rendering."""

import re
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta

EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+(?:\.[\w-]+)+$")
OFFSET_RE = re.compile(
    r"^\s*(?P<n>\d+)\s*(?P<unit>m|mins?|minutes?|h|hrs?|hours?|d|days?|w|wks?|weeks?)"
    r"\s*(?P<dir>before|after)?\s*$",
    re.I,
)
UNIT_NAMES = {"m": "minute", "h": "hour", "d": "day", "w": "week"}


@dataclass(frozen=True)
class Offset:
    n: int
    unit: str  # m, h, d, w
    after: bool = False

    @property
    def token(self):
        return f"{self.n}{self.unit}{' after' if self.after else ''}"

    @property
    def delta(self):
        return {
            "m": timedelta(minutes=self.n),
            "h": timedelta(hours=self.n),
            "d": timedelta(days=self.n),
            "w": timedelta(weeks=self.n),
        }[self.unit]

    @property
    def label(self):
        if self.n == 0:
            return "day of" if self.unit in "dw" else "at start"
        name = UNIT_NAMES[self.unit] + ("s" if self.n != 1 else "")
        return f"{self.n} {name} {'after' if self.after else 'before'}"


def parse_offsets(text):
    """Parse '14d, 7d, 2h, 14d after' into Offsets. Returns (offsets, errors)."""
    offsets, errors = [], []
    for part in re.split(r"[,;\n]+", text or ""):
        if not part.strip():
            continue
        m = OFFSET_RE.match(part)
        if not m:
            errors.append(part.strip())
            continue
        unit = m.group("unit")[0].lower()
        off = Offset(int(m.group("n")), unit, (m.group("dir") or "").lower() == "after")
        if off not in offsets:
            offsets.append(off)
    return offsets, errors


def parse_recipients(text):
    """Split a comma/semicolon/space separated list. Returns (valid, invalid)."""
    valid, invalid = [], []
    for part in re.split(r"[,;\s]+", text or ""):
        part = part.strip().strip("<>")
        if not part:
            continue
        if EMAIL_RE.match(part):
            if part.lower() not in (v.lower() for v in valid):
                valid.append(part)
        else:
            invalid.append(part)
    return valid, invalid


def _parse_hhmm(value, fallback=time(9, 0)):
    try:
        return datetime.strptime(value.strip(), "%H:%M").time()
    except (AttributeError, ValueError):
        return fallback


def event_start(event, all_day_send_time="09:00"):
    d = date.fromisoformat(event["start_date"])
    t = _parse_hhmm(event["start_time"]) if event["start_time"] else _parse_hhmm(all_day_send_time)
    return datetime.combine(d, t)


def event_end(event, all_day_send_time="09:00"):
    d = date.fromisoformat(event["end_date"]) if event["end_date"] else date.fromisoformat(event["start_date"])
    if event["end_time"]:
        t = _parse_hhmm(event["end_time"])
    elif event["start_time"]:
        t = _parse_hhmm(event["start_time"])
    else:
        t = _parse_hhmm(all_day_send_time)
    return datetime.combine(d, t)


def compute_send_time(event, offset, all_day_send_time="09:00"):
    anchor = event_end(event, all_day_send_time) if offset.after else event_start(event, all_day_send_time)
    return anchor + offset.delta if offset.after else anchor - offset.delta


def planned_reminders(event, all_day_send_time="09:00"):
    """List of (Offset, send_at datetime) sorted by send time."""
    offsets, _ = parse_offsets(event["offsets"])
    plans = [(o, compute_send_time(event, o, all_day_send_time)) for o in offsets]
    return sorted(plans, key=lambda p: p[1])


def time_until_phrase(offset):
    if offset.after:
        if offset.n == 0:
            return "just ended"
        return f"{offset.n} {UNIT_NAMES[offset.unit]}{'s' if offset.n != 1 else ''} ago"
    if offset.n == 0:
        return "today" if offset.unit in "dw" else "starting now"
    if offset.unit == "d" and offset.n == 1:
        return "tomorrow"
    if offset.unit == "d" and offset.n % 7 == 0:
        weeks = offset.n // 7
        return f"in {weeks} week{'s' if weeks != 1 else ''}"
    return f"in {offset.n} {UNIT_NAMES[offset.unit]}{'s' if offset.n != 1 else ''}"


def _fmt_date(d):
    return f"{d:%A, %B} {d.day}, {d.year}"


def _fmt_time(t):
    return f"{t.hour % 12 or 12}:{t.minute:02d} {'AM' if t.hour < 12 else 'PM'}"


def describe_when(event):
    start = date.fromisoformat(event["start_date"])
    end = date.fromisoformat(event["end_date"]) if event["end_date"] else None
    st = _parse_hhmm(event["start_time"]) if event["start_time"] else None
    et = _parse_hhmm(event["end_time"]) if event["end_time"] else None

    text = _fmt_date(start) + (f" at {_fmt_time(st)}" if st else "")
    if end and end != start:
        text += f" – {_fmt_date(end)}" + (f" at {_fmt_time(et)}" if et else "")
    elif et:
        text += f" – {_fmt_time(et)}"
    return text


def render_email(event, offset):
    title = event["title"]
    time_until = time_until_phrase(offset)
    if offset.after:
        reminder_line = f"This is a follow-up reminder for {title}, which ended {time_until}."
    else:
        reminder_line = f"This is a reminder that {title} is starting {time_until}."

    details = [
        f"- Course: {title}",
        f"- Program: {event['program'] or 'N/A'}",
        f"- Start Date: {event['start_date']}",
        f"- End Date: {event['end_date'] or 'N/A'}",
    ]
    if event["start_time"]:
        details.append(f"- Time: {_fmt_time(_parse_hhmm(event['start_time']))}")

    body = "Dear Team,\n\n" + reminder_line + "\n\nCourse Details:\n" + "\n".join(details) + "\n"
    return f"Course Reminder: {title}", body
