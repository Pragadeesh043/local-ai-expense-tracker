"""Monthly TOTAL rows, Dashboard sheet and chart, and the summarize command."""

import re
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from typer.testing import CliRunner

from expense_tracker import cli
from expense_tracker.chart_builder import CHART_TITLE, DASHBOARD, category_totals
from expense_tracker.schema import Expense
from expense_tracker.sheet import (
    SheetError,
    append_expense,
    append_expenses,
    read_recent,
    remove_totals,
    summarize,
)


def exp(d, amount=100, cat="Other", cur="INR", **kw):
    return Expense(date=d, category=cat, amount=amount, currency=cur, raw_note="n", **kw)


def seed(p: Path, order="desc"):
    """Two October rows and three November rows."""
    append_expenses(
        p, "Expenses",
        [
            exp("2026-10-10", 100, "Transport"),
            exp("2026-10-20", 200, "Food & Dining"),
            exp("2026-11-02", 300, "Food & Dining"),
            exp("2026-11-05", 400, "Groceries"),
            exp("2026-11-08", 50, "Transport"),
        ],
        order=order,
    )


def col_a(p: Path):
    ws = load_workbook(p)["Expenses"]
    return [ws.cell(r, 1).value for r in range(2, ws.max_row + 1)]


def labels(p: Path):
    return [v if isinstance(v, str) else v.strftime("%Y-%m-%d") for v in col_a(p)]


# ---------- monthly totals ----------


def test_no_totals_until_summarize_is_run(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    assert not any(isinstance(v, str) for v in col_a(p))


def test_totals_desc_layout_and_formulas(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    res = summarize(p, "Expenses", order="desc", dashboard=False)
    assert res.months == 2 and res.totals_on
    assert labels(p) == [
        "2026-11-08", "2026-11-05", "2026-11-02", "TOTAL (2026-11)",
        "2026-10-20", "2026-10-10", "TOTAL (2026-10)",
    ]
    ws = load_workbook(p)["Expenses"]
    assert ws["C5"].value == "=SUM(C2:C4)"  # November block is rows 2-4, total on row 5
    assert ws["C8"].value == "=SUM(C6:C7)"  # October block is rows 6-7, total on row 8
    assert ws["D5"].value == "INR"


def test_totals_asc_layout_and_formulas(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p, order="asc")
    summarize(p, "Expenses", order="asc", dashboard=False)
    assert labels(p) == [
        "2026-10-10", "2026-10-20", "TOTAL (2026-10)",
        "2026-11-02", "2026-11-05", "2026-11-08", "TOTAL (2026-11)",
    ]
    ws = load_workbook(p)["Expenses"]
    assert ws["C4"].value == "=SUM(C2:C3)"
    assert ws["C8"].value == "=SUM(C5:C7)"


def test_total_rows_are_highlighted_across_the_row(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    summarize(p, "Expenses", dashboard=False)
    ws = load_workbook(p)["Expenses"]
    for col in "ABCDEFG":
        cell = ws[f"{col}5"]
        assert cell.font.bold is True
        assert cell.fill.fgColor.rgb.endswith("FFF2CC")
    assert ws["C5"].number_format == "#,##0.00"
    # a normal row is not highlighted
    assert ws["A2"].font.bold is not True
    assert not ws["A2"].fill.fgColor.rgb.endswith("FFF2CC")
    # header untouched
    assert ws["A1"].font.bold is True and ws.freeze_panes == "A2"


def test_totals_are_idempotent(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    summarize(p, "Expenses", dashboard=False)
    first = labels(p)
    summarize(p, "Expenses", dashboard=False)
    summarize(p, "Expenses", dashboard=False)
    assert labels(p) == first


def test_add_after_summarize_keeps_and_updates_totals(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    summarize(p, "Expenses", dashboard=False)

    r = append_expense(p, "Expenses", exp("2026-11-09", 10))  # newest row goes to the top
    assert r == 2
    ws = load_workbook(p)["Expenses"]
    assert ws["A6"].value == "TOTAL (2026-11)"
    assert ws["C6"].value == "=SUM(C2:C5)"  # range grew by one row
    assert ws["A9"].value == "TOTAL (2026-10)"

    append_expense(p, "Expenses", exp("2026-12-01", 5))  # a new month gets its own total
    assert "TOTAL (2026-12)" in labels(p)
    assert labels(p)[:2] == ["2026-12-01", "TOTAL (2026-12)"]


def test_add_into_older_month_lands_in_that_block(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    summarize(p, "Expenses", dashboard=False)
    r = append_expense(p, "Expenses", exp("2026-10-15", 7))
    ws = load_workbook(p)["Expenses"]
    assert ws.cell(r, 3).value == 7
    # October block now has 3 rows and its total covers them
    month_rows = [i for i in range(2, ws.max_row + 1) if ws.cell(i, 1).value == "TOTAL (2026-10)"]
    assert ws.cell(month_rows[0], 3).value == f"=SUM(C{month_rows[0]-3}:C{month_rows[0]-1})"


def test_remove_totals(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    summarize(p, "Expenses", dashboard=False)
    res = remove_totals(p, "Expenses")
    assert res.months == 0 and not res.totals_on
    assert labels(p) == ["2026-11-08", "2026-11-05", "2026-11-02", "2026-10-20", "2026-10-10"]
    ws = load_workbook(p)["Expenses"]
    assert ws.max_row == 6  # leftover rows deleted
    append_expense(p, "Expenses", exp("2026-11-09", 1))  # and they do not come back
    assert not any("TOTAL" in str(v) for v in col_a(p))


def test_multi_currency_month_gets_one_total_per_currency(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expenses(p, "Expenses", [exp("2026-11-02", 100), exp("2026-11-03", 20, cur="USD"), exp("2026-11-04", 50)])
    summarize(p, "Expenses", dashboard=False)
    ws = load_workbook(p)["Expenses"]
    totals = [(ws.cell(r, 3).value, ws.cell(r, 4).value) for r in range(2, ws.max_row + 1) if str(ws.cell(r, 1).value).startswith("TOTAL")]
    assert totals == [
        ('=SUMIF(D2:D4,"INR",C2:C4)', "INR"),
        ('=SUMIF(D2:D4,"USD",C2:C4)', "USD"),
    ]


def test_rows_without_a_readable_date_stay_last_with_no_total(tmp_path):
    p = tmp_path / "e.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    ws.append(["Date", "Category", "Amount", "Currency", "Merchant", "Payment Method", "Raw Note"])
    ws.append(["???", "Other", 9, "INR"])
    wb.save(p)
    append_expense(p, "Expenses", exp("2026-11-02", 1))
    summarize(p, "Expenses", dashboard=False)
    assert labels(p) == ["2026-11-02", "TOTAL (2026-11)", "???"]


def test_read_recent_skips_total_rows(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    summarize(p, "Expenses", dashboard=False)
    rows = read_recent(p, "Expenses", 20)
    assert len(rows) == 5
    assert not any(str(r[0]).startswith("TOTAL") for r in rows)


def test_workbook_recalculates_on_open(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    summarize(p, "Expenses")
    assert load_workbook(p).calculation.fullCalcOnLoad is True


def test_summarize_missing_file(tmp_path):
    with pytest.raises(SheetError, match="does not exist"):
        summarize(tmp_path / "nope.xlsx", "Expenses")


# ---------- dashboard ----------


def chart_title(ch):
    return ch.title.tx.rich.p[0].r[0].t


def test_dashboard_table_and_chart(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    res = summarize(p, "Expenses")
    assert res.dashboard_written and res.categories == 3
    wb = load_workbook(p)
    assert DASHBOARD in wb.sheetnames and wb.sheetnames[0] == "Expenses"
    ws = wb[DASHBOARD]
    # largest first: Food & Dining 500, Groceries 400, Transport 150
    assert [ws.cell(r, 1).value for r in (5, 6, 7)] == ["Food & Dining", "Groceries", "Transport"]
    assert ws["B5"].value == "=SUMIFS('Expenses'!$C:$C,'Expenses'!$B:$B,$A5,'Expenses'!$D:$D,\"INR\")"
    assert ws["A8"].value == "Total" and ws["B8"].value == "=SUM(B5:B7)"
    assert len(ws._charts) == 1
    ch = ws._charts[0]
    assert chart_title(ch) == CHART_TITLE == "Spending Breakdown by Category"
    assert type(ch).__name__ == "BarChart"
    ref = ch.series[0].val.numRef.f
    assert ref == "'Dashboard'!$B$4:$B$7" or ref == "Dashboard!$B$4:$B$7" or ref.endswith("$B$5:$B$7")


def test_dashboard_is_refreshed_on_add_without_duplicating_the_chart(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    summarize(p, "Expenses")
    append_expense(p, "Expenses", exp("2026-11-09", 9, "Health"))
    ws = load_workbook(p)[DASHBOARD]
    cats = [ws.cell(r, 1).value for r in range(5, 12) if ws.cell(r, 1).value not in (None, "Total")]
    assert "Health" in cats
    assert len(ws._charts) == 1


def test_no_dashboard_flag_and_no_dashboard_unless_asked(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    assert DASHBOARD not in load_workbook(p).sheetnames
    summarize(p, "Expenses", dashboard=False)
    assert DASHBOARD not in load_workbook(p).sheetnames


def test_other_sheets_survive(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    wb = load_workbook(p)
    wb.create_sheet("Budget")["A1"] = "keep me"
    wb.save(p)
    summarize(p, "Expenses")
    append_expense(p, "Expenses", exp("2026-11-09", 1))
    assert load_workbook(p)["Budget"]["A1"].value == "keep me"


def test_dashboard_other_currency_note_and_scope(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expenses(p, "Expenses", [exp("2026-11-02", 100, "Transport"), exp("2026-11-03", 20, "Transport", cur="USD")])
    summarize(p, "Expenses", currency="INR")
    ws = load_workbook(p)[DASHBOARD]
    assert ws["B5"].value.endswith('"INR")')
    notes = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("Not included")]
    assert notes and "USD (1 rows)" in notes[0]


def test_dashboard_with_no_matching_expenses_has_no_chart(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", exp("2026-11-02", 20, cur="USD"))
    res = summarize(p, "Expenses", currency="INR")
    assert res.categories == 0
    ws = load_workbook(p)[DASHBOARD]
    assert ws["A5"].value == "No expenses in INR yet."
    assert len(ws._charts) == 0


def test_category_totals_helper():
    rows = [("Food", 10, "INR"), ("Food", 5.5, "inr"), ("Fuel", 20, None), ("Food", 9, "USD"), ("", 5, "INR"), ("Fuel", "x", "INR")]
    ordered, other = category_totals(rows, "INR")
    assert ordered == [("Fuel", 20.0), ("Food", 15.5)]
    assert other == {"USD": 1}


# ---------- CLI ----------


def run(*args):
    return CliRunner().invoke(cli.app, list(args))


def test_cli_summarize_end_to_end(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    r = run("summarize", "-f", str(p))
    assert r.exit_code == 0, r.output
    out = " ".join(r.output.split())
    assert "2 monthly total row group(s)" in out
    assert "chart of 3 categories" in out
    wb = load_workbook(p)
    assert DASHBOARD in wb.sheetnames
    assert any(str(v).startswith("TOTAL") for v in col_a(p))


def test_cli_summarize_no_dashboard_and_remove(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    assert run("summarize", "-f", str(p), "--no-dashboard").exit_code == 0
    assert DASHBOARD not in load_workbook(p).sheetnames
    r = run("summarize", "-f", str(p), "--remove")
    assert r.exit_code == 0 and "Removed" in r.output
    assert not any(str(v).startswith("TOTAL") for v in col_a(p))


def test_cli_summarize_missing_file(tmp_path):
    r = run("summarize", "-f", str(tmp_path / "nope.xlsx"))
    assert r.exit_code == 1 and "does not exist" in r.output


def test_cli_list_ignores_total_rows(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    summarize(p, "Expenses")
    r = run("list", "-f", str(p), "-n", "20")
    assert r.exit_code == 0, r.output
    assert "TOTAL" not in r.output
