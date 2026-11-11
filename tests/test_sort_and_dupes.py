"""Chronological sorting on append, and protection against duplicate rows."""

from copy import copy
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from typer.testing import CliRunner

from expense_tracker import cli, sheet
from expense_tracker.config import Settings
from expense_tracker.extract import extract_expenses
from expense_tracker.schema import Expense
from functools import partial

from expense_tracker.sheet import HEADERS, read_recent
from expense_tracker.sheet import append_expense as _append_expense
from expense_tracker.sheet import append_expenses as _append_expenses

# These tests are about row order, so they write without monthly TOTAL rows.
append_expense = partial(_append_expense, totals=False)
append_expenses = partial(_append_expenses, totals=False)

TODAY = date(2026, 11, 9)


def exp(d, amount=100, cat="OTHERS", merchant="N/A", **kw):
    return Expense(date=d, category=cat, amount=amount, merchant=merchant, raw_note="n", **kw)


def dates_in(p: Path, sheet_name="Expenses"):
    ws = load_workbook(p)[sheet_name]
    return [ws.cell(r, 1).value.strftime("%Y-%m-%d") for r in range(2, ws.max_row + 1)]


def amounts_in(p: Path):
    ws = load_workbook(p)["Expenses"]
    return [ws.cell(r, 3).value for r in range(2, ws.max_row + 1)]


# ---------- sorting ----------


def test_older_entry_added_later_goes_below_newer_desc(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", exp("2026-11-08", 1))
    append_expense(p, "Expenses", exp("2026-11-06", 2))  # last Friday, added after yesterday
    append_expense(p, "Expenses", exp("2026-11-09", 3))
    assert dates_in(p) == ["2026-11-09", "2026-11-08", "2026-11-06"]
    assert amounts_in(p) == [3, 1, 2]


def test_ascending_order(tmp_path):
    p = tmp_path / "e.xlsx"
    for d, a in [("2026-11-08", 1), ("2026-11-06", 2), ("2026-11-09", 3)]:
        append_expense(p, "Expenses", exp(d, a), order="asc")
    assert dates_in(p) == ["2026-11-06", "2026-11-08", "2026-11-09"]


def test_returned_rows_are_final_positions(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", exp("2026-11-08", 1))
    rows = append_expenses(p, "Expenses", [exp("2026-11-01", 2), exp("2026-11-09", 3)])
    assert rows == [4, 2]  # final order: 11-09 (row 2), 11-08 (row 3), 11-01 (row 4)
    ws = load_workbook(p)["Expenses"]
    assert [ws.cell(r, 3).value for r in rows] == [2, 3]


def test_same_date_keeps_insertion_order(tmp_path):
    p = tmp_path / "e.xlsx"
    for a in (1, 2, 3):
        append_expense(p, "Expenses", exp("2026-11-08", a))
    assert amounts_in(p) == [1, 2, 3]
    append_expenses(p, "Expenses", [exp("2026-11-08", 4), exp("2026-11-08", 5)])
    assert amounts_in(p) == [1, 2, 3, 4, 5]


def test_existing_unsorted_rows_get_sorted(tmp_path):
    p = tmp_path / "e.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    for i, h in enumerate(HEADERS, 1):
        ws.cell(1, i, h)
    for r, (d, a) in enumerate([(date(2026, 11, 1), 10), (date(2026, 11, 7), 20), (date(2026, 11, 3), 30)], 2):
        ws.cell(r, 1, d)
        ws.cell(r, 3, a)
    wb.save(p)
    append_expense(p, "Expenses", exp("2026-11-05", 40))
    assert dates_in(p) == ["2026-11-07", "2026-11-05", "2026-11-03", "2026-11-01"]
    assert amounts_in(p) == [20, 40, 30, 10]


def test_header_and_formatting_survive_resort(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", exp("2026-11-08", 1))
    append_expense(p, "Expenses", exp("2026-11-09", 2))
    ws = load_workbook(p)["Expenses"]
    assert [c.value for c in ws[1]] == HEADERS
    assert ws["A1"].font.bold is True
    assert ws["A1"].fill.fgColor.rgb.endswith("1F4E78")
    assert ws.freeze_panes == "A2"
    assert ws.column_dimensions["G"].width == 50
    assert ws["A2"].number_format == "YYYY-MM-DD"
    assert ws["C2"].number_format == "#,##0.00"


def test_row_styles_move_with_their_row(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", exp("2026-11-01", 1))
    wb = load_workbook(p)
    ws = wb["Expenses"]
    ws["A2"].fill = PatternFill("solid", fgColor="FFFF00")  # highlight the old row
    wb.save(p)

    append_expense(p, "Expenses", exp("2026-11-09", 2))  # pushes the highlighted row to row 3
    ws = load_workbook(p)["Expenses"]
    assert ws["A3"].fill.fgColor.rgb.endswith("FFFF00")
    assert ws["C3"].value == 1
    assert ws["A2"].fill.fgColor.rgb != ws["A3"].fill.fgColor.rgb  # new row has no highlight


def test_formulas_follow_their_row(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", exp("2026-11-01", 10))
    wb = load_workbook(p)
    ws = wb["Expenses"]
    ws["H2"] = "=C2*2"
    wb.save(p)

    append_expense(p, "Expenses", exp("2026-11-09", 5))  # old row moves from 2 to 3
    ws = load_workbook(p)["Expenses"]
    assert ws["C3"].value == 10
    assert ws["H3"].value == "=C3*2"


def test_unreadable_dates_go_last_and_blank_rows_dropped(tmp_path):
    p = tmp_path / "e.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    for i, h in enumerate(HEADERS, 1):
        ws.cell(1, i, h)
    ws.cell(2, 1, "not a date")
    ws.cell(2, 3, 7)
    # row 3 blank
    ws.cell(4, 1, datetime(2026, 11, 2))
    ws.cell(4, 3, 8)
    wb.save(p)

    append_expense(p, "Expenses", exp("2026-11-05", 9))
    ws = load_workbook(p)["Expenses"]
    assert ws.max_row == 4
    assert [ws.cell(r, 3).value for r in (2, 3, 4)] == [9, 8, 7]


def test_one_save_for_many_rows(tmp_path, monkeypatch):
    p = tmp_path / "e.xlsx"
    from openpyxl.workbook.workbook import Workbook as WB

    calls = []
    real = WB.save
    monkeypatch.setattr(WB, "save", lambda self, f: (calls.append(f), real(self, f))[1])
    append_expenses(p, "Expenses", [exp("2026-11-08", 1), exp("2026-11-07", 2), exp("2026-11-06", 3)])
    assert len(calls) == 1


def test_locked_file_writes_nothing_and_keeps_order(tmp_path, monkeypatch):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", exp("2026-11-08", 1))
    before = p.read_bytes()
    from openpyxl.workbook.workbook import Workbook as WB

    def boom(self, f):
        raise PermissionError("locked")

    monkeypatch.setattr(WB, "save", boom)
    with pytest.raises(sheet.SheetError, match="Close it in Excel"):
        append_expense(p, "Expenses", exp("2026-11-09", 2))
    assert p.read_bytes() == before


def test_read_recent_uses_date_not_position(tmp_path):
    p = tmp_path / "e.xlsx"
    for d, a in [("2026-11-08", 1), ("2026-11-01", 2), ("2026-11-09", 3), ("2026-11-05", 4)]:
        append_expense(p, "Expenses", exp(d, a))
    assert [r[2] for r in read_recent(p, "Expenses", 2)] == [1, 3]  # the two newest, oldest first


# ---------- duplicates ----------


class FakeClient:
    def __init__(self, *responses):
        self.responses = list(responses)

    def chat_json(self, system, user, schema):
        return self.responses.pop(0)


def S(tmp_path, **kw):
    return Settings(xlsx_path=tmp_path / "e.xlsx", **kw)


def it(cat, amount, **kw):
    return {"category": cat, "amount": amount, **kw}


def test_model_repeating_an_item_is_collapsed(tmp_path):
    c = FakeClient({"expenses": [it("FOOD", 700, merchant="Zomato"), it("FOOD", 700, merchant="Zomato")]})
    out = extract_expenses("zomato 700", S(tmp_path), client=c, today=TODAY)
    assert len(out) == 1 and out[0].amount == 700.0


def test_real_repeats_are_kept(tmp_path):
    c = FakeClient({"expenses": [it("FOOD", 50), it("FOOD", 50)]})
    out = extract_expenses("coffee 50 + tea 50", S(tmp_path), client=c, today=TODAY)
    assert len(out) == 2


def test_triple_from_model_but_two_in_note(tmp_path):
    c = FakeClient({"expenses": [it("OTHERS", 50)] * 3})
    assert len(extract_expenses("a 50 b 50", S(tmp_path), client=c, today=TODAY)) == 2


def test_repeat_check_handles_commas_and_decimals(tmp_path):
    c = FakeClient({"expenses": [it("OTHERS", 1250.5)] * 2})
    assert len(extract_expenses("rent 1,250.50", S(tmp_path), client=c, today=TODAY)) == 1


def test_different_items_with_different_amounts_untouched(tmp_path):
    c = FakeClient({"expenses": [it("TRANSPORT", 490), it("FOOD", 330)]})
    assert len(extract_expenses("cab 490 + lunch 330", S(tmp_path), client=c, today=TODAY)) == 2


def test_cli_writes_each_parsed_item_exactly_once(tmp_path, monkeypatch):
    p = tmp_path / "e.xlsx"
    calls = []
    real = cli.append_expenses
    monkeypatch.setattr(cli, "append_expenses", lambda *a, **k: (calls.append(a), real(*a, **k))[1])

    class C:  # model that doubles its answer
        def chat_json(self, *_):
            return {"expenses": [it("FOOD", 700, merchant="Zomato")] * 2}

    monkeypatch.setattr(cli, "OllamaClient", lambda *a, **k: C())
    r = CliRunner().invoke(cli.app, ["add", "-y", "-f", str(p), "zomato", "700"])
    assert r.exit_code == 0, r.output
    assert len(calls) == 1
    ws = load_workbook(p)["Expenses"]
    assert ws.max_row == 3  # header, ONE data row, and its monthly TOTAL row
    assert ws['C2'].value == 700
