"""Automatic monthly TOTAL rows and the Dashboard (KPI cards, per-month tables and pie charts)."""

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from typer.testing import CliRunner

from expense_tracker import cli
from expense_tracker.config import Settings
from expense_tracker.dashboard_builder import CATEGORY_COLORS, DASHBOARD, month_tables
from expense_tracker.income import Income
from expense_tracker.schema import Expense
from expense_tracker.sheet import (
    append_expense,
    append_expenses,
    append_incomes,
    read_recent,
    summarize,
)


def exp(d, amount=100, cat="OTHERS", cur="INR", **kw):
    return Expense(date=d, category=cat, amount=amount, currency=cur, raw_note="n", **kw)


def seed(p: Path, order="desc", totals=True):
    """Two October rows and three November rows."""
    append_expenses(
        p, "Expenses",
        [
            exp("2026-10-10", 100, "TRANSPORT"),
            exp("2026-10-20", 200, "FOOD"),
            exp("2026-11-02", 300, "FOOD"),
            exp("2026-11-05", 400, "GROCERY"),
            exp("2026-11-08", 50, "TRANSPORT"),
        ],
        order=order, totals=totals,
    )


def col_a(p: Path):
    ws = load_workbook(p)["Expenses"]
    return [ws.cell(r, 1).value for r in range(2, ws.max_row + 1)]


def labels(p: Path):
    return [v if isinstance(v, str) else v.strftime("%Y-%m-%d") for v in col_a(p)]


# ---------- monthly totals ----------


def test_totals_are_added_automatically_desc(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)  # no summarize needed
    assert labels(p) == [
        "2026-11-08", "2026-11-05", "2026-11-02", "TOTAL (2026-11)",
        "2026-10-20", "2026-10-10", "TOTAL (2026-10)",
    ]
    ws = load_workbook(p)["Expenses"]
    assert ws["C5"].value == "=SUM(C2:C4)"
    assert ws["C8"].value == "=SUM(C6:C7)"
    assert ws["D5"].value == "INR"


def test_totals_asc(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p, order="asc")
    assert labels(p) == [
        "2026-10-10", "2026-10-20", "TOTAL (2026-10)",
        "2026-11-02", "2026-11-05", "2026-11-08", "TOTAL (2026-11)",
    ]
    ws = load_workbook(p)["Expenses"]
    assert ws["C4"].value == "=SUM(C2:C3)" and ws["C8"].value == "=SUM(C5:C7)"


def test_total_rows_bold_gray_across_the_row(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    ws = load_workbook(p)["Expenses"]
    for col in "ABCDEFG":
        cell = ws[f"{col}5"]
        assert cell.font.bold is True
        assert cell.fill.fgColor.rgb.endswith("EDEDED")
    assert ws["C5"].number_format == "#,##0.00"
    assert ws["A2"].font.bold is not True
    assert ws["A1"].font.bold is True and ws.freeze_panes == "A2"


def test_add_keeps_totals_correct(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    assert append_expense(p, "Expenses", exp("2026-11-09", 10)) == 2
    ws = load_workbook(p)["Expenses"]
    assert ws["A6"].value == "TOTAL (2026-11)" and ws["C6"].value == "=SUM(C2:C5)"
    assert ws["A9"].value == "TOTAL (2026-10)"
    append_expense(p, "Expenses", exp("2026-12-01", 5))
    assert labels(p)[:2] == ["2026-12-01", "TOTAL (2026-12)"]
    assert labels(p).count("TOTAL (2026-11)") == 1  # rebuilt, never duplicated


def test_totals_can_be_switched_off(tmp_path, monkeypatch):
    p = tmp_path / "e.xlsx"
    seed(p)
    summarize(p, "Expenses", totals=False)
    assert not any("TOTAL" in str(v) for v in col_a(p))
    assert load_workbook(p)["Expenses"].max_row == 6  # leftover rows removed
    monkeypatch.setenv("EXPENSE_TOTALS", "0")
    assert Settings().monthly_totals is False
    monkeypatch.setenv("EXPENSE_TOTALS", "1")
    assert Settings().monthly_totals is True
    monkeypatch.delenv("EXPENSE_TOTALS")
    assert Settings().monthly_totals is True


def test_summarize_is_idempotent(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    first = labels(p)
    summarize(p, "Expenses")
    summarize(p, "Expenses")
    assert labels(p) == first


def test_multi_currency_month_gets_one_total_per_currency(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expenses(p, "Expenses", [exp("2026-11-02", 100), exp("2026-11-03", 20, cur="USD"), exp("2026-11-04", 50)])
    ws = load_workbook(p)["Expenses"]
    totals = [(ws.cell(r, 3).value, ws.cell(r, 4).value) for r in range(2, ws.max_row + 1) if str(ws.cell(r, 1).value).startswith("TOTAL")]
    assert totals == [('=SUMIF(D2:D4,"INR",C2:C4)', "INR"), ('=SUMIF(D2:D4,"USD",C2:C4)', "USD")]


def test_unreadable_dates_stay_last_without_a_total(tmp_path):
    p = tmp_path / "e.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    ws.append(["Date", "Category", "Amount", "Currency", "Merchant", "Payment Method", "Raw Note"])
    ws.append(["???", "OTHERS", 9, "INR"])
    wb.save(p)
    append_expense(p, "Expenses", exp("2026-11-02", 1))
    assert labels(p) == ["2026-11-02", "TOTAL (2026-11)", "???"]


def test_read_recent_skips_total_rows(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    rows = read_recent(p, "Expenses", 20)
    assert len(rows) == 5 and not any(str(r[0]).startswith("TOTAL") for r in rows)


def test_old_category_labels_are_migrated(tmp_path):
    p = tmp_path / "e.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    ws.append(["Date", "Category", "Amount", "Currency", "Merchant", "Payment Method", "Raw Note"])
    for i, c in enumerate(["Food & Dining", "Groceries", "Utilities", "Shopping", "Health", "Subscriptions", "Transport"]):
        ws.append([f"2026-11-0{i + 1}", c, 10, "INR"])
    wb.save(p)
    append_expense(p, "Expenses", exp("2026-11-09", 1))
    ws = load_workbook(p)["Expenses"]
    cats = [ws.cell(r, 2).value for r in range(2, ws.max_row + 1) if ws.cell(r, 2).value]
    assert sorted(cats) == sorted(["OTHERS", "FOOD", "GROCERY", "OTHERS", "PERSONAL", "PERSONAL", "ENTERTAINMENT", "TRANSPORT"])


def test_workbook_recalculates_on_open(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    assert load_workbook(p).calculation.fullCalcOnLoad is True


def test_summarize_missing_file(tmp_path):
    from expense_tracker.sheet import SheetError

    with pytest.raises(SheetError, match="does not exist"):
        summarize(tmp_path / "nope.xlsx", "Expenses")


# ---------- dashboard ----------


def dash(p: Path):
    return load_workbook(p)[DASHBOARD]


def chart_title(ch):
    return ch.title.tx.rich.p[0].r[0].t


def test_dashboard_is_first_and_active(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    wb = load_workbook(p)
    assert wb.sheetnames == ["Dashboard", "Expenses", "Income"]
    assert wb.active.title == "Dashboard"
    assert [s.sheet_view.tabSelected for s in wb.worksheets] == [True, False, False]


def test_kpi_cards(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    ws = dash(p)
    assert ws["A4"].value == "TOTAL INCOME (INR)" and ws["A5"].value == "=SUM('Income'!$C:$C)"
    assert ws["E4"].value == "TOTAL SPENT (INR)"
    # The monthly TOTAL rows have no category, so this formula never counts them twice.
    assert ws["E5"].value == "=SUMIFS('Expenses'!$C:$C,'Expenses'!$B:$B,\"<>\",'Expenses'!$D:$D,\"INR\")"
    assert ws["I4"].value == "NET BALANCE (INR)" and ws["I5"].value == "=A5-E5"
    merged = {str(r) for r in ws.merged_cells.ranges}
    assert {"A4:C4", "A5:C5", "E4:G4", "E5:G5", "I4:K4", "I5:K5"} <= merged
    assert ws["A5"].font.bold and ws["A5"].font.size == 22
    assert ws["A4"].border.top.style and ws["A5"].border.bottom.style and ws["C5"].border.right.style
    assert len(ws.conditional_formatting) >= 1  # net balance turns red when negative


def test_one_table_and_one_pie_per_month(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    ws = dash(p)
    assert ws["A8"].value == "November 2026" and ws["A25"].value == "October 2026"
    # November: FOOD 300, GROCERY 400, TRANSPORT 50 -> sorted by spend, all 7 categories listed
    assert [ws.cell(r, 1).value for r in range(10, 17)] == [
        "GROCERY", "FOOD", "TRANSPORT", "PARKING", "ENTERTAINMENT", "PERSONAL", "OTHERS",
    ]
    assert ws["A17"].value == "Total" and ws["B17"].value == "=SUM(B10:B16)"
    assert ws["B10"].value == (
        "=SUMIFS('Expenses'!$C:$C,'Expenses'!$B:$B,$A10,"
        "'Expenses'!$A:$A,\">=\"&DATE(2026,11,1),'Expenses'!$A:$A,\"<\"&DATE(2026,12,1),"
        "'Expenses'!$D:$D,\"INR\")"
    )
    assert ws["C10"].value == "=IF($B$17=0,0,B10/$B$17)"
    assert len(ws._charts) == 2
    titles = [chart_title(c) for c in ws._charts]
    assert titles == ["November 2026: Spending by Category", "October 2026: Spending by Category"]
    assert all(type(c).__name__ == "PieChart" for c in ws._charts)


def test_pie_covers_only_categories_with_spending_and_uses_bright_colors(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    nov = dash(p)._charts[0]
    ser = nov.series[0]
    assert ser.val.numRef.f.endswith("$B$10:$B$12")  # the 3 spending rows, not the 4 zero rows
    assert ser.tx.strRef.f.endswith("B9")  # series name comes from the table header
    assert (ser.cat.strRef or ser.cat.numRef).f.endswith("$A$10:$A$12")
    colors = [pt.graphicalProperties.solidFill.srgbClr for pt in ser.dPt]
    assert colors == [CATEGORY_COLORS["GROCERY"], CATEGORY_COLORS["FOOD"], CATEGORY_COLORS["TRANSPORT"]]
    assert len(set(CATEGORY_COLORS.values())) == 7  # every category has its own color
    assert nov.dataLabels.showPercent is True


def test_charts_sit_beside_their_tables(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    a, b = (c.anchor._from for c in dash(p)._charts)
    assert (a.col, a.row) == (4, 8)  # E9, next to the November table header
    assert (b.col, b.row) == (4, 25)  # E26, next to the October table header


def test_income_and_spending_formulas_use_the_right_sheets(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    append_incomes(p, "Expenses", [Income(date="2026-11-01", source="Salary", amount=75000)])
    ws = dash(p)
    assert ws["A5"].value == "=SUM('Income'!$C:$C)"
    assert len(ws._charts) == 2  # still two, not duplicated


def test_dashboard_is_rebuilt_on_add_without_duplicating(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    append_expense(p, "Expenses", exp("2026-11-09", 9, "PARKING"))
    append_expense(p, "Expenses", exp("2026-09-01", 9, "PERSONAL"))
    wb = load_workbook(p)
    assert wb.sheetnames.count(DASHBOARD) == 1
    ws = wb[DASHBOARD]
    assert len(ws._charts) == 3
    assert [ws.cell(r, 1).value for r in (8, 25, 42)] == ["November 2026", "October 2026", "September 2026"]


def test_december_rolls_into_next_year(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", exp("2026-12-05", 9, "FOOD"))
    assert 'DATE(2027,1,1)' in dash(p)["B10"].value


def test_month_with_no_inr_spending_has_no_chart(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expenses(p, "Expenses", [exp("2026-11-02", 100, "FOOD"), exp("2026-09-03", 20, "FOOD", cur="USD")])
    ws = dash(p)
    assert len(ws._charts) == 1
    assert ws["E27"].value == "No spending in INR this month."
    assert "USD (1 rows)" in ws["A6"].value


def test_empty_workbook_dashboard(tmp_path):
    p = tmp_path / "e.xlsx"
    append_incomes(p, "Expenses", [Income(date="2026-11-01", source="Salary", amount=5)])
    ws = dash(p)
    assert ws["A8"].value == "No expenses yet." and len(ws._charts) == 0
    assert ws["A5"].value == "=SUM('Income'!$C:$C)"


def test_other_sheets_survive(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    wb = load_workbook(p)
    wb.create_sheet("Budget")["A1"] = "keep me"
    wb.save(p)
    append_expense(p, "Expenses", exp("2026-11-09", 1))
    wb = load_workbook(p)
    assert wb["Budget"]["A1"].value == "keep me" and wb.sheetnames[0] == "Dashboard"


def test_month_tables_helper():
    rows = [("2026-11-02", "FOOD", 10, "INR"), ("2026-11-03", "Food & Dining", 5, None), ("2026-10-01", "PARKING", 7, "INR"), ("2026-11-04", "FOOD", 3, "USD"), ("bad", "FOOD", 1, "INR")]
    months, other = month_tables(rows, "INR")
    assert list(months) == ["2026-11", "2026-10"]
    assert months["2026-11"]["FOOD"] == 15.0 and months["2026-10"]["PARKING"] == 7.0
    assert other == {"USD": 1}


# ---------- CLI ----------


def run(*args):
    return CliRunner().invoke(cli.app, list(args))


def test_cli_summarize(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p, totals=False)
    assert not any("TOTAL" in str(v) for v in col_a(p))
    r = run("summarize", "-f", str(p))
    assert r.exit_code == 0, r.output
    out = " ".join(r.output.split())
    assert "monthly TOTAL rows" in out and "2 monthly chart(s)" in out
    assert any(str(v).startswith("TOTAL") for v in col_a(p))


def test_cli_summarize_respects_totals_off(tmp_path, monkeypatch):
    p = tmp_path / "e.xlsx"
    seed(p)
    monkeypatch.setenv("EXPENSE_TOTALS", "0")
    r = run("summarize", "-f", str(p))
    assert r.exit_code == 0 and "EXPENSE_TOTALS is off" in " ".join(r.output.split())
    assert not any("TOTAL" in str(v) for v in col_a(p))


def test_cli_summarize_missing_file(tmp_path):
    r = run("summarize", "-f", str(tmp_path / "nope.xlsx"))
    assert r.exit_code == 1 and "does not exist" in r.output


def test_cli_list_ignores_total_rows(tmp_path):
    p = tmp_path / "e.xlsx"
    seed(p)
    r = run("list", "-f", str(p), "-n", "20")
    assert r.exit_code == 0 and "TOTAL" not in r.output
