"""Turn a schedule pasted in (almost) any format, or an uploaded spreadsheet, into draft events.

Free text is parsed line by line with date/time patterns. Tables (tab-separated text copied
from Excel, CSV, or .xlsx uploads) are mapped by column headers. If OPENAI_API_KEY is set,
free text is sent to an LLM first and the pattern parser is used as a fallback.
"""

import csv
import io
import json
import os
import re
import urllib.request
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, time, timedelta

import pandas as pd


@dataclass
class ParsedEvent:
    title: str
    start_date: str
    start_time: str = ""
    end_date: str = ""
    end_time: str = ""
    location: str = ""
    notes: str = ""
    recipients: str = ""

    def to_dict(self):
        return asdict(self)


@dataclass
class ParseResult:
    events: list = field(default_factory=list)
    skipped: list = field(default_factory=list)
    method: str = "patterns"


MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
WEEKDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

MONTH = (
    r"(?P<mon>jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?"
    r"|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?"
)
WD_PREFIX = (
    r"(?:(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday"
    r"|mon|tues?|wed|thurs?|thu|fri|sat|sun)\.?,?\s+)?"
)
ORD = r"(?:st|nd|rd|th)?"
YEAR = r"(?:,?\s+(?P<y>\d{4}))?"
NOT_TIME = r"(?!:\d)(?!\s*[ap]\.?m\b)"
DASH = r"(?:-|–|—|to|until|till|through|thru)"

DATE_PATTERNS = [
    ("iso", re.compile(r"\b(?P<y>\d{4})-(?P<m>\d{1,2})-(?P<d>\d{1,2})(?:T(?=\d))?", re.I)),
    ("md_range", re.compile(
        rf"\b{WD_PREFIX}{MONTH}\s+(?P<d>\d{{1,2}}){ORD}\s*{DASH}\s*(?P<d2>\d{{1,2}}){ORD}\b{NOT_TIME}{YEAR}",
        re.I)),
    ("md", re.compile(rf"\b{WD_PREFIX}{MONTH}\s+(?P<d>\d{{1,2}}){ORD}\b{NOT_TIME}{YEAR}", re.I)),
    ("dm", re.compile(rf"\b{WD_PREFIX}(?P<d>\d{{1,2}}){ORD}\s+(?:of\s+)?{MONTH}\b{YEAR}", re.I)),
    ("dm_sep", re.compile(rf"\b(?P<d>\d{{1,2}})[-/.]{MONTH}(?:[-/.](?P<y>\d{{4}}|\d{{2}}))?\b", re.I)),
    ("md_sep", re.compile(rf"\b{MONTH}[-/.](?P<d>\d{{1,2}})(?:[-/.](?P<y>\d{{4}}|\d{{2}}))?\b", re.I)),
    ("numeric", re.compile(r"\b(?P<a>\d{1,2})/(?P<b>\d{1,2})(?:/(?P<y>\d{2,4}))?\b")),
    ("relative", re.compile(r"\b(?P<rel>today|tonight|tomorrow)\b", re.I)),
    ("weekday", re.compile(
        r"\b(?P<next>next\s+|this\s+)?(?P<wd>monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b",
        re.I)),
]

AMPM = r"(?P<ap{n}>[ap])\.?m\.?(?![a-z])"
TIME_PATTERNS = [
    ("range12", re.compile(
        r"\b(?P<h1>\d{1,2})(?::(?P<m1>\d{2}))?\s*(?:" + AMPM.format(n=1) + r")?\s*" + DASH
        + r"\s*(?P<h2>\d{1,2})(?::(?P<m2>\d{2}))?\s*" + AMPM.format(n=2), re.I)),
    ("range24", re.compile(
        r"\b(?P<h1>[01]?\d|2[0-3]):(?P<m1>[0-5]\d)\s*" + DASH
        + r"\s*(?P<h2>[01]?\d|2[0-3]):(?P<m2>[0-5]\d)\b", re.I)),
    ("t12", re.compile(r"\b(?P<h1>\d{1,2})(?::(?P<m1>\d{2}))?\s*" + AMPM.format(n=1), re.I)),
    ("t24", re.compile(r"\b(?P<h1>[01]?\d|2[0-3]):(?P<m1>[0-5]\d)\b")),
    ("word", re.compile(r"\b(?P<word>noon|midday|midnight)\b", re.I)),
]

EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
FIELD_LINE_RE = re.compile(
    r"^(?P<key>location|where|room|venue|notes?|details?|description|recipients?|attendees|emails?|to)"
    r"\s*:\s*(?P<value>.+)$",
    re.I,
)
BULLET_RE = re.compile(r"^\s*(?:[-*•▪◦·]+|\d{1,3}[.)])\s+")
CONNECTOR_BEFORE = re.compile(r"(?:\b(?:on|at|from|between|by|starting|starts)\b|@)\s*\x00", re.I)
JOIN_SPANS = re.compile(r"\x00\s*(?:,|-|–|—|\b(?:to|until|till|and|at|from)\b)?\s*\x00", re.I)
LOCATION_RE = re.compile(
    r"\s(?:(?:@|\b(?:location|where|room|venue)\s*:)\s*(?P<loc>[^,;|]+)"
    r"|[-–—,;|]\s*(?P<room>(?:room|rm\.?|bldg\.?|building)\s+[^,;|]+))$",
    re.I,
)


def _resolve_date(y, m, d, today):
    if y:
        y = int(y)
        if y < 100:
            y += 2000
        return date(y, m, d)
    candidate = date(today.year, m, d)
    if candidate < today - timedelta(days=30):
        candidate = date(today.year + 1, m, d)
    return candidate


def _date_from_match(kind, m, today):
    """Return (start_date, end_date_or_None) for a date pattern match, or None if invalid."""
    g = m.groupdict()
    try:
        if kind == "iso":
            return date(int(g["y"]), int(g["m"]), int(g["d"])), None
        if kind in ("md", "dm", "md_range", "dm_sep", "md_sep"):
            month = MONTHS[g["mon"][:3].lower()]
            start = _resolve_date(g.get("y"), month, int(g["d"]), today)
            if kind == "md_range":
                d2 = int(g["d2"])
                if d2 >= start.day:
                    end = start.replace(day=d2)
                else:
                    nm = start.replace(day=1) + timedelta(days=32)
                    end = nm.replace(day=d2)
                return start, end
            return start, None
        if kind == "numeric":
            a, b = int(g["a"]), int(g["b"])
            month, day = (b, a) if a > 12 else (a, b)
            return _resolve_date(g.get("y"), month, day, today), None
        if kind == "relative":
            rel = g["rel"].lower()
            return (today + timedelta(days=1) if rel == "tomorrow" else today), None
        if kind == "weekday":
            target = WEEKDAYS.index(g["wd"].lower())
            delta = (target - today.weekday()) % 7
            if g.get("next") and g["next"].strip().lower() == "next" and delta == 0:
                delta = 7
            return today + timedelta(days=delta), None
    except (ValueError, KeyError):
        return None
    return None


def _to_24h(h, mins, ap):
    h, mins = int(h), int(mins or 0)
    if ap:
        if h < 1 or h > 12:
            return None
        h = h % 12 + (12 if ap.lower() == "p" else 0)
    elif h > 23:
        return None
    if mins > 59:
        return None
    return time(h, mins)


def _time_from_match(kind, m):
    """Return (start_time, end_time_or_None) for a time match, or None if invalid."""
    g = m.groupdict()
    if kind == "word":
        return (time(0, 0) if g["word"].lower() == "midnight" else time(12, 0)), None
    if kind in ("t12", "t24"):
        t = _to_24h(g["h1"], g["m1"], g.get("ap1"))
        return (t, None) if t else None
    if kind == "range24":
        t1, t2 = _to_24h(g["h1"], g["m1"], None), _to_24h(g["h2"], g["m2"], None)
        return (t1, t2) if t1 and t2 else None
    if kind == "range12":
        t2 = _to_24h(g["h2"], g["m2"], g["ap2"])
        if not t2:
            return None
        if g.get("ap1"):
            t1 = _to_24h(g["h1"], g["m1"], g["ap1"])
        else:
            t1 = _to_24h(g["h1"], g["m1"], g["ap2"])
            if t1 and t1 > t2:
                t1 = _to_24h(g["h1"], g["m1"], "a" if g["ap2"].lower() == "p" else "p")
        return (t1, t2) if t1 else None
    return None


def _find_spans(text, patterns, convert):
    """Find non-overlapping matches, preferring longer ones. Returns sorted list of (start, end, value)."""
    candidates = []
    for kind, pattern in patterns:
        for m in pattern.finditer(text):
            value = convert(kind, m)
            if value:
                candidates.append((m.start(), m.end(), value))
    candidates.sort(key=lambda c: (-(c[1] - c[0]), c[0]))
    chosen = []
    for c in candidates:
        if all(c[1] <= o[0] or c[0] >= o[1] for o in chosen):
            chosen.append(c)
    return sorted(chosen)


def _mask(text, spans):
    chars = list(text)
    for start, end, _ in spans:
        chars[start:end] = ["\x00"] + [" "] * (end - start - 1)
    return "".join(chars)


def _clean_title(masked):
    s = masked
    for _ in range(3):
        s = JOIN_SPANS.sub("\x00", s)
        s = CONNECTOR_BEFORE.sub("\x00", s)
    s = s.replace("\x00", " ")
    s = re.sub(r"\(\s*\)", " ", s)
    s = re.sub(r"\s+([,;:])", r"\1", s)
    s = re.sub(r"(?:\s*[-–—|,;:]\s*){2,}", " - ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip(" -–—|,;:@\t")


def parse_line(line, today):
    """Parse one line. Returns dict with dates, times, title, emails (any may be empty)."""
    emails = EMAIL_RE.findall(line)
    text = EMAIL_RE.sub(" ", line)
    date_spans = _find_spans(text, DATE_PATTERNS, lambda k, m: _date_from_match(k, m, today))
    masked = _mask(text, date_spans)
    time_spans = _find_spans(masked, TIME_PATTERNS, _time_from_match)
    masked = _mask(masked, time_spans)
    title = _clean_title(masked)
    location = ""
    loc = LOCATION_RE.search(" " + title)
    if loc:
        location = (loc.group("loc") or loc.group("room")).strip()
        title = _clean_title((" " + title)[: loc.start()])

    dates = []
    for _, _, (start, end) in date_spans:
        dates.append(start)
        if end:
            dates.append(end)
    times = []
    for _, _, (start, end) in time_spans:
        times.append(start)
        if end:
            times.append(end)
    return {"dates": dates, "times": times, "title": title, "location": location, "emails": emails}


def _make_event(title, dates, times, location="", emails=None):
    start = dates[0]
    end = dates[-1] if len(dates) > 1 and dates[-1] != start else None
    ev = ParsedEvent(
        title=title or "Untitled event",
        start_date=start.isoformat(),
        start_time=times[0].strftime("%H:%M") if times else "",
        end_date=end.isoformat() if end else "",
        location=location,
        recipients=", ".join(emails or []),
    )
    if len(times) > 1:
        ev.end_time = times[1].strftime("%H:%M")
        if not ev.end_date and times[1] < times[0]:
            ev.end_date = (start + timedelta(days=1)).isoformat()
    return ev


def _apply_field_line(ev, key, value):
    key = key.lower()
    if key in ("location", "where", "room", "venue"):
        ev.location = value
    elif key.startswith(("recipient", "attendee", "email")) or key == "to":
        found = EMAIL_RE.findall(value)
        if found:
            ev.recipients = ", ".join(filter(None, [ev.recipients] + found))
    else:
        ev.notes = (ev.notes + "\n" + value).strip()


def parse_free_text(text, today=None):
    today = today or date.today()
    result = ParseResult()
    context_date = None
    pending_title = None
    last_event = None

    for raw in text.splitlines():
        line = BULLET_RE.sub("", raw).strip()
        if not line:
            continue

        field_match = FIELD_LINE_RE.match(line)
        if field_match and last_event and not parse_line(field_match.group("value"), today)["dates"]:
            _apply_field_line(last_event, field_match.group("key"), field_match.group("value").strip())
            continue

        p = parse_line(line, today)
        has_title = bool(p["title"])

        if p["dates"]:
            if has_title:
                if pending_title:
                    result.skipped.append(pending_title)
                    pending_title = None
                last_event = _make_event(p["title"], p["dates"], p["times"], p["location"], p["emails"])
                result.events.append(last_event)
            elif pending_title:
                last_event = _make_event(pending_title, p["dates"], p["times"], p["location"], p["emails"])
                result.events.append(last_event)
                pending_title = None
            else:
                context_date = p["dates"][0]
                if p["times"]:
                    last_event = None
            continue

        if p["times"] and has_title and context_date:
            if pending_title:
                result.skipped.append(pending_title)
                pending_title = None
            last_event = _make_event(p["title"], [context_date], p["times"], p["location"], p["emails"])
            result.events.append(last_event)
            continue

        if pending_title:
            result.skipped.append(pending_title)
        pending_title = line

    if pending_title:
        result.skipped.append(pending_title)
    return result


# --- Tables ---------------------------------------------------------------------------------

COLUMN_KEYWORDS = {
    "title": ["course name", "event name", "session name", "title", "event", "course", "subject",
              "session", "task", "activity", "meeting", "item", "name"],
    "start_time": ["start time", "time", "from time", "begins at"],
    "end_time": ["end time", "to time", "finish time", "ends at"],
    "start_date": ["start date", "start", "date", "when", "begin", "from", "day", "due"],
    "end_date": ["end date", "end", "finish", "until", "to", "through"],
    "location": ["location", "room", "where", "venue", "place", "building"],
    "recipients": ["recipients", "emails", "email", "attendees", "invitees", "send to"],
    "notes": ["notes", "note", "details", "comments", "description"],
}


def _match_columns(headers):
    lowered = {h: str(h).strip().lower() for h in headers}
    mapping = {}
    used = set()
    for fieldname in ["start_time", "end_time", "start_date", "end_date", "title", "location",
                      "recipients", "notes"]:
        for kw in COLUMN_KEYWORDS[fieldname]:
            hit = next(
                (h for h, low in lowered.items() if h not in used
                 and (low == kw or re.search(rf"\b{re.escape(kw)}\b", low))),
                None,
            )
            if hit is not None:
                if fieldname == "start_date" and "time" in lowered[hit]:
                    continue
                mapping[fieldname] = hit
                used.add(hit)
                break
    return mapping


def _cell_to_date_time(value, today):
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None, None
    if isinstance(value, (pd.Timestamp, datetime)):
        t = value.time()
        return value.date(), (t if t != time(0, 0) else None)
    if isinstance(value, date):
        return value, None
    if isinstance(value, time):
        return None, value
    p = parse_line(str(value), today)
    return (p["dates"][0] if p["dates"] else None), (p["times"][0] if p["times"] else None)


def _cell_text(value):
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def parse_rows(df, today=None, sheet_name=None):
    today = today or date.today()
    result = ParseResult(method="table")
    df = df.dropna(how="all").dropna(axis=1, how="all")
    mapping = _match_columns(list(df.columns))
    if "start_date" not in mapping:
        result.skipped.append(
            f"{'Sheet ' + repr(sheet_name) + ': ' if sheet_name else ''}"
            f"couldn't find a date column (columns: {', '.join(map(str, df.columns))})"
        )
        return result
    extra_cols = [c for c in df.columns if c not in mapping.values() and not str(c).startswith("Unnamed")]

    for _, row in df.iterrows():
        start_d, start_t = _cell_to_date_time(row[mapping["start_date"]], today)
        title = _cell_text(row[mapping["title"]]) if "title" in mapping else ""
        if not start_d:
            text = " | ".join(_cell_text(v) for v in row.values if _cell_text(v))
            if text:
                result.skipped.append(text)
            continue
        end_d, end_t = (None, None)
        if "end_date" in mapping:
            end_d, end_t = _cell_to_date_time(row[mapping["end_date"]], today)
        if "start_time" in mapping:
            _, t = _cell_to_date_time(row[mapping["start_time"]], today)
            start_t = t or start_t
        if "end_time" in mapping:
            _, t = _cell_to_date_time(row[mapping["end_time"]], today)
            end_t = t or end_t

        notes = [_cell_text(row[mapping["notes"]])] if "notes" in mapping else []
        notes += [f"{c}: {_cell_text(row[c])}" for c in extra_cols if _cell_text(row[c])]
        if sheet_name:
            notes.append(f"Sheet: {sheet_name}")
        recipients = EMAIL_RE.findall(_cell_text(row[mapping["recipients"]])) if "recipients" in mapping else []

        result.events.append(ParsedEvent(
            title=title or "Untitled event",
            start_date=start_d.isoformat(),
            start_time=start_t.strftime("%H:%M") if start_t else "",
            end_date=end_d.isoformat() if end_d and end_d != start_d else "",
            end_time=end_t.strftime("%H:%M") if end_t else "",
            location=_cell_text(row[mapping["location"]]) if "location" in mapping else "",
            notes="\n".join(n for n in notes if n),
            recipients=", ".join(recipients),
        ))
    return result


def _looks_tabular(text):
    lines = [l for l in text.splitlines() if l.strip()]
    if len(lines) < 2:
        return None
    if sum("\t" in l for l in lines) >= max(2, len(lines) * 0.6):
        return "\t"
    header = lines[0].lower()
    if header.count(",") >= 2 and re.search(r"\b(date|start|when|day)\b", header):
        return ","
    return None


def parse_upload(filename, data, today=None):
    name = filename.lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        sheets = pd.read_excel(io.BytesIO(data), sheet_name=None)
        result = ParseResult(method="spreadsheet")
        for sheet, df in sheets.items():
            sub = parse_rows(df, today, sheet_name=sheet if len(sheets) > 1 else None)
            result.events += sub.events
            result.skipped += sub.skipped
        return result
    text = data.decode("utf-8-sig", errors="replace")
    if name.endswith(".csv"):
        return parse_rows(pd.read_csv(io.StringIO(text)), today)
    return parse_text(text, today)


# --- Optional LLM parsing ---------------------------------------------------------------------

LLM_PROMPT = """Extract every dated event, deadline or session from the schedule below.
Today is {today}. Resolve relative dates ("next Tuesday") and missing years to the next upcoming date.
Return JSON: {{"events": [{{"title": str, "start_date": "YYYY-MM-DD", "start_time": "HH:MM" or "",
"end_date": "YYYY-MM-DD" or "", "end_time": "HH:MM" or "", "location": str, "notes": str,
"recipients": [email, ...]}}], "skipped": [lines you could not turn into an event]}}.
Use 24-hour times. Leave start_time empty for all-day items. Do not invent events.

Schedule:
{text}"""


def parse_with_llm(text, today):
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return None
    body = json.dumps({
        "model": os.environ.get("OPENAI_MODEL", "gpt-4o-mini"),
        "response_format": {"type": "json_object"},
        "messages": [{"role": "user", "content": LLM_PROMPT.format(today=today.isoformat(), text=text)}],
    }).encode()
    req = urllib.request.Request(
        os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1") + "/chat/completions",
        data=body,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        content = json.load(resp)["choices"][0]["message"]["content"]
    data = json.loads(content)
    result = ParseResult(method="ai", skipped=[str(s) for s in data.get("skipped", [])])
    for e in data.get("events", []):
        try:
            date.fromisoformat(e["start_date"])
        except (KeyError, TypeError, ValueError):
            result.skipped.append(str(e.get("title", e)))
            continue
        result.events.append(ParsedEvent(
            title=str(e.get("title") or "Untitled event"),
            start_date=e["start_date"],
            start_time=e.get("start_time") or "",
            end_date=e.get("end_date") or "",
            end_time=e.get("end_time") or "",
            location=e.get("location") or "",
            notes=e.get("notes") or "",
            recipients=", ".join(r for r in e.get("recipients") or [] if EMAIL_RE.fullmatch(r)),
        ))
    return result


def parse_text(text, today=None):
    today = today or date.today()
    sep = _looks_tabular(text)
    if sep:
        rows = list(csv.reader(io.StringIO(text.strip()), delimiter=sep))
        width = max(len(r) for r in rows)
        rows = [r + [""] * (width - len(r)) for r in rows]
        df = pd.DataFrame(rows[1:], columns=[c.strip() or f"Unnamed {i}" for i, c in enumerate(rows[0])])
        df = df.replace("", None)
        return parse_rows(df, today)
    return parse_free_text(text, today)


def parse_schedule(text, today=None, use_ai=True):
    """Parse pasted text. Returns (ParseResult, warning_or_None)."""
    today = today or date.today()
    if use_ai and os.environ.get("OPENAI_API_KEY") and not _looks_tabular(text):
        try:
            return parse_with_llm(text, today), None
        except Exception as exc:
            return parse_text(text, today), f"AI parsing failed ({exc}); used the built-in parser instead."
    return parse_text(text, today), None
