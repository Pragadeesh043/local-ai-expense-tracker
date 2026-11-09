from datetime import date

import pytest

from expense_tracker.dates import resolve_date

TODAY = date(2026, 11, 5)  # Thursday


@pytest.mark.parametrize(
    "text,expected",
    [
        ("", TODAY),
        (None, TODAY),
        ("today", TODAY),
        ("yesterday", date(2026, 11, 4)),
        ("day before yesterday", date(2026, 11, 3)),
        ("3 days ago", date(2026, 11, 2)),
        ("2 weeks ago", date(2026, 10, 22)),
        ("last friday", date(2026, 10, 30)),
        ("Friday", date(2026, 10, 30)),
        ("last thursday", date(2026, 10, 29)),  # same weekday as today -> a week back
        ("3 nov", date(2026, 11, 3)),
        ("3rd November", date(2026, 11, 3)),
        ("nov 3", date(2026, 11, 3)),
        ("15 dec", date(2025, 12, 15)),  # future month/day -> previous year
        ("2026-10-01", date(2026, 10, 1)),
        ("03/11/2026", date(2026, 11, 3)),
        ("3/11", date(2026, 11, 3)),
        ("garbage words", TODAY),
        ("31 feb", TODAY),
    ],
)
def test_resolve_date(text, expected):
    assert resolve_date(text, TODAY) == expected
