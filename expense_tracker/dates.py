"""Deterministic date resolution.

Small models are unreliable at date arithmetic, so the LLM only copies the date phrase it
sees. This module turns that phrase into an ISO date. It also scans the original note
itself, so a missed or wrong phrase from the model does not decide the date.
"""

from __future__ import annotations

import re
from datetime import date, timedelta

_WEEKDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
_MONTHS = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}
_WD = "|".join(_WEEKDAYS)
_MON = "|".join(sorted(_MONTHS, key=len, reverse=True))

# Patterns used to find a date phrase inside a full note. Ordered by priority.
# Bare weekdays ("friday") and short numeric dates ("3/11") are NOT matched here:
# they collide with merchant names and amounts. The model's own phrase may use them.
_NOTE_PATTERNS = [
    r"\bday before yesterday\b",
    r"\b(?:yesterday|last night)\b",
    r"\b\d+\s+(?:days?|weeks?)\s+(?:ago|back)\b",
    r"\b(?:last week|a week ago)\b",
    rf"\b(?:last|this past|past|on)\s+(?:{_WD})\b",
    r"\b\d{4}-\d{1,2}-\d{1,2}\b",
    rf"\b\d{{1,2}}(?:st|nd|rd|th)?\s+(?:{_MON})\b(?:\s+\d{{4}})?",
    rf"\b(?:{_MON})\s+\d{{1,2}}(?:st|nd|rd|th)?\b(?:,?\s+\d{{4}})?",
    r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",
]


def _last_weekday(today: date, target: int) -> date:
    """Most recent past occurrence of weekday `target` (strictly before today)."""
    delta = (today.weekday() - target) % 7
    return today - timedelta(days=delta or 7)


def _safe(y: int, m: int, d: int) -> date | None:
    try:
        return date(y, m, d)
    except ValueError:
        return None


def _month_day(today: date, month: int, day: int, year: str | None) -> date | None:
    if year:
        return _safe(int(year), month, day)
    d = _safe(today.year, month, day)
    if d is None:
        return None
    # No year given and the date is in the future -> assume the previous year.
    return d if d <= today else _safe(today.year - 1, month, day)


def parse_date_phrase(text: str | None, today: date) -> date | None:
    """Resolve a date phrase. Returns None when nothing recognizable is found."""
    if not text:
        return None
    t = text.strip().lower()
    if not t:
        return None

    if re.search(r"\bday before yesterday\b", t):
        return today - timedelta(days=2)
    if re.search(r"\b(?:yesterday|last night)\b", t):
        return today - timedelta(days=1)

    m = re.search(r"\b(\d+)\s+days?\s+(?:ago|back)\b", t)
    if m:
        return today - timedelta(days=int(m.group(1)))
    m = re.search(r"\b(\d+)\s+weeks?\s+(?:ago|back)\b", t)
    if m:
        return today - timedelta(weeks=int(m.group(1)))
    if re.search(r"\b(?:last week|a week ago)\b", t):
        return today - timedelta(weeks=1)

    m = re.search(rf"\b(?:last|this past|past|on)\s+({_WD})\b", t)
    if m:
        return _last_weekday(today, _WEEKDAYS.index(m.group(1)))
    if t in _WEEKDAYS:
        return _last_weekday(today, _WEEKDAYS.index(t))

    m = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", t)
    if m:
        return _safe(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    # 3 nov / 3rd november / 3 nov 2026
    m = re.search(rf"\b(\d{{1,2}})(?:st|nd|rd|th)?\s+({_MON})\b(?:\s+(\d{{4}}))?", t)
    if m:
        return _month_day(today, _MONTHS[m.group(2)], int(m.group(1)), m.group(3))

    # nov 3 / november 3rd / nov 3, 2026
    m = re.search(rf"\b({_MON})\s+(\d{{1,2}})(?:st|nd|rd|th)?\b(?:,?\s+(\d{{4}}))?", t)
    if m:
        return _month_day(today, _MONTHS[m.group(1)], int(m.group(2)), m.group(3))

    # 03/11/2026 or 3/11 -> day first (India)
    m = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?", t)
    if m:
        year = m.group(3)
        if year and len(year) == 2:
            year = "20" + year
        return _month_day(today, int(m.group(2)), int(m.group(1)), year)

    if re.search(r"\b(?:today|tonight|this morning|this afternoon|this evening|now)\b", t):
        return today
    return None


def resolve_date(text: str | None, today: date) -> date:
    """Like parse_date_phrase, but falls back to `today`."""
    return parse_date_phrase(text, today) or today


def find_date_in_note(note: str, today: date) -> date | None:
    """Find the first clear date phrase inside the user's own note."""
    n = note.lower()
    for pattern in _NOTE_PATTERNS:
        m = re.search(pattern, n)
        if m:
            d = parse_date_phrase(m.group(0), today)
            if d:
                return d
    return None


def resolve_item_date(model_phrase: str | None, note: str, today: date) -> date:
    """Final date for one expense.

    1. Use the model's phrase only if it really appears in the note. A phrase the user never
       wrote (for example "today" for a note that says "yesterday") is a model mistake.
    2. Otherwise scan the note ourselves.
    3. Otherwise use today.
    """
    phrase = (model_phrase or "").strip().lower()
    if phrase and phrase in note.lower():
        d = parse_date_phrase(phrase, today)
        if d:
            return d
    return find_date_in_note(note, today) or today


def cell_date(v) -> date | None:
    """Best-effort date from a spreadsheet cell value (datetime, date or ISO text). None if not a date."""
    from datetime import datetime

    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        try:
            return date.fromisoformat(v.strip()[:10])
        except ValueError:
            return None
    return None
