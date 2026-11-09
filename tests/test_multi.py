"""Multi-item notes and date handling when the model gets dates wrong."""

from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook
from typer.testing import CliRunner

from expense_tracker import cli
from expense_tracker.config import Settings
from expense_tracker.dates import find_date_in_note, parse_date_phrase, resolve_item_date
from expense_tracker.extract import ExtractionError, extract_expenses
from expense_tracker.schema import MAX_ITEMS, Expense
from expense_tracker.sheet import append_expenses

TODAY = date(2026, 11, 9)  # Monday


class FakeClient:
    def __init__(self, *responses):
        self.responses = list(responses)
        self.calls = []

    def chat_json(self, system, user, schema):
        self.calls.append((system, user, schema))
        return self.responses.pop(0)


def S(tmp_path: Path, **kw) -> Settings:
    return Settings(xlsx_path=tmp_path / "e.xlsx", **kw)


def item(cat, amount, **kw):
    return {"category": cat, "amount": amount, **kw}


# ---------- multi-item ----------


def test_two_items_two_rows(tmp_path):
    c = FakeClient(
        {
            "expenses": [
                item("Transport", 490, payment_method="Cash"),
                item("Food & Dining", 330, payment_method="Cash"),
            ]
        }
    )
    out = extract_expenses("cab 490 + lunch 330 paid via cash", S(tmp_path), client=c, today=TODAY)
    assert [(e.category, e.amount, e.payment_method) for e in out] == [
        ("Transport", 490.0, "Cash"),
        ("Food & Dining", 330.0, "Cash"),
    ]
    assert all(e.raw_note == "cab 490 + lunch 330 paid via cash" for e in out)


def test_bare_list_and_single_object_still_accepted(tmp_path):
    c = FakeClient([item("Other", 1), item("Other", 2)], item("Other", 3))
    assert len(extract_expenses("a 1 b 2", S(tmp_path), client=c, today=TODAY)) == 2
    assert len(extract_expenses("c 3", S(tmp_path), client=c, today=TODAY)) == 1


def test_empty_list_is_retried_then_fails(tmp_path):
    c = FakeClient({"expenses": []}, {"expenses": []}, {"expenses": []})
    with pytest.raises(ExtractionError):
        extract_expenses("x 1", S(tmp_path), client=c, today=TODAY)
    assert len(c.calls) == 3


def test_one_bad_item_retries_whole_note(tmp_path):
    bad = {"expenses": [item("Other", 1), item("Other", "oops")]}
    good = {"expenses": [item("Other", 1), item("Other", 2)]}
    c = FakeClient(bad, good)
    out = extract_expenses("a 1 b 2", S(tmp_path), client=c, today=TODAY)
    assert [e.amount for e in out] == [1.0, 2.0]


def test_too_many_items_rejected(tmp_path):
    many = {"expenses": [item("Other", 1)] * (MAX_ITEMS + 1)}
    c = FakeClient(many, many, many)
    with pytest.raises(ExtractionError, match="limit"):
        extract_expenses("x", S(tmp_path), client=c, today=TODAY)


def test_prompt_demands_one_object_per_item(tmp_path):
    c = FakeClient({"expenses": [item("Other", 1)]})
    extract_expenses("x 1", S(tmp_path), client=c, today=TODAY)
    system, _, schema = c.calls[0]
    assert "TWO objects" in system and "Never add amounts together" in system
    assert "cab 490 + lunch 330" in system  # worked example
    assert schema["properties"]["expenses"]["type"] == "array"


# ---------- dates ----------


@pytest.mark.parametrize(
    "text,expected",
    [
        ("yesterday evening", date(2026, 11, 8)),
        ("last night", date(2026, 11, 8)),
        ("the day before yesterday", date(2026, 11, 7)),
        ("3 days back", date(2026, 11, 6)),
        ("on friday", date(2026, 11, 6)),
        ("last monday", date(2026, 11, 2)),  # same weekday as today -> a week back
        ("this evening", TODAY),
    ],
)
def test_parse_embedded_phrases(text, expected):
    assert parse_date_phrase(text, TODAY) == expected


def test_parse_unknown_returns_none():
    assert parse_date_phrase("some words", TODAY) is None
    assert parse_date_phrase("", TODAY) is None
    assert parse_date_phrase(None, TODAY) is None


@pytest.mark.parametrize(
    "note,expected",
    [
        ("swiggy 450 yesterday", date(2026, 11, 8)),
        ("cab 490 + lunch 330 paid via cash yesterday", date(2026, 11, 8)),
        ("uber 200 2 days ago", date(2026, 11, 7)),
        ("dinner 900 on 3 nov", date(2026, 11, 3)),
        ("dinner 900 nov 3rd", date(2026, 11, 3)),
        ("fuel 500 last friday", date(2026, 11, 6)),
        ("rent 9000 2026-10-01", date(2026, 10, 1)),
    ],
)
def test_find_date_in_note(note, expected):
    assert find_date_in_note(note, TODAY) == expected


@pytest.mark.parametrize(
    "note",
    [
        "Friday's Pizza 300",  # bare weekday inside a merchant name is not a date
        "cab 490 + lunch 330",
        "sent 3/4 of the bill 300",  # short numeric is not scanned in notes
    ],
)
def test_find_date_in_note_ignores_false_positives(note):
    assert find_date_in_note(note, TODAY) is None


def test_model_phrase_not_in_note_is_ignored():
    # Model says "today" but the user wrote "yesterday": the note wins.
    assert resolve_item_date("today", "swiggy 450 yesterday", TODAY) == date(2026, 11, 8)


def test_model_empty_phrase_falls_back_to_note_scan():
    assert resolve_item_date("", "swiggy 450 yesterday", TODAY) == date(2026, 11, 8)
    assert resolve_item_date(None, "swiggy 450 yesterday", TODAY) == date(2026, 11, 8)


def test_unparseable_model_phrase_falls_back_to_note_scan():
    assert resolve_item_date("sometime", "lunch sometime, bought yesterday", TODAY) == date(2026, 11, 8)


def test_no_date_anywhere_is_today():
    assert resolve_item_date("", "swiggy 450", TODAY) == TODAY


def test_model_bare_weekday_phrase_works_when_in_note():
    assert resolve_item_date("friday", "gym 800 friday", TODAY) == date(2026, 11, 6)


def test_yesterday_end_to_end_with_model_leaving_date_blank(tmp_path):
    c = FakeClient({"expenses": [item("Food & Dining", 450, date_text="")]})
    [e] = extract_expenses("swiggy 450 yesterday", S(tmp_path), client=c, today=TODAY)
    assert e.date == "2026-11-08"


def test_yesterday_end_to_end_with_model_saying_today(tmp_path):
    c = FakeClient({"expenses": [item("Food & Dining", 450, date_text="today")]})
    [e] = extract_expenses("swiggy 450 yesterday", S(tmp_path), client=c, today=TODAY)
    assert e.date == "2026-11-08"


def test_per_item_dates_when_note_has_two(tmp_path):
    c = FakeClient(
        {
            "expenses": [
                item("Transport", 490, date_text="yesterday"),
                item("Food & Dining", 330, date_text="today"),
            ]
        }
    )
    out = extract_expenses("cab 490 yesterday + lunch 330 today", S(tmp_path), client=c, today=TODAY)
    assert [e.date for e in out] == ["2026-11-08", "2026-11-09"]


def test_shared_date_applies_to_every_item_when_model_omits_it(tmp_path):
    c = FakeClient({"expenses": [item("Transport", 490), item("Food & Dining", 330)]})
    out = extract_expenses("cab 490 + lunch 330 yesterday", S(tmp_path), client=c, today=TODAY)
    assert [e.date for e in out] == ["2026-11-08", "2026-11-08"]


# ---------- sheet ----------


def _exp(**kw):
    base = dict(date="2026-11-08", category="Transport", amount=490, raw_note="n")
    base.update(kw)
    return Expense(**base)


def test_append_expenses_writes_consecutive_rows(tmp_path):
    p = tmp_path / "e.xlsx"
    rows = append_expenses(p, "Expenses", [_exp(), _exp(category="Food & Dining", amount=330)])
    assert rows == [2, 3]
    more = append_expenses(p, "Expenses", [_exp(amount=1)])
    assert more == [4]
    ws = load_workbook(p)["Expenses"]
    assert [ws.cell(r, 3).value for r in (2, 3, 4)] == [490, 330, 1]
    assert ws["B3"].value == "Food & Dining"


def test_append_expenses_empty_writes_nothing(tmp_path):
    p = tmp_path / "e.xlsx"
    assert append_expenses(p, "Expenses", []) == []
    assert not p.exists()


# ---------- CLI ----------


def test_cli_add_multi_item(tmp_path, monkeypatch):
    p = tmp_path / "e.xlsx"
    items = [_exp(), _exp(category="Food & Dining", amount=330, payment_method="Cash")]
    monkeypatch.setattr(cli, "extract_expenses", lambda note, s, client=None: items)
    r = CliRunner().invoke(cli.app, ["add", "-y", "-f", str(p), "cab", "490", "+", "lunch", "330"])
    assert r.exit_code == 0, r.output
    assert "Parsed 2 expenses" in r.output
    assert "rows 2, 3" in " ".join(r.output.split())
    ws = load_workbook(p)["Expenses"]
    assert ws.max_row == 3
