"""Dashboard sheet: spend per category (live formulas) and a bar chart.

The sheet is generated. It is deleted and rebuilt on every refresh, so do not put your own
work on it.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Iterable

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook import Workbook

DASHBOARD = "Dashboard"
CHART_TITLE = "Spending Breakdown by Category"

_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
_TOP = Border(top=Side(style="thin"))
_FIRST_ROW = 5  # first category row; the table header is row 4


def _quote(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"


def category_totals(
    rows: Iterable[tuple], currency: str
) -> tuple[list[tuple[str, float]], dict[str, int]]:
    """Total per category in `currency`, largest first, plus a count of rows in other currencies.

    rows: (category, amount, currency) tuples. Rows with a blank category or a non-numeric
    amount are skipped.
    """
    cur = currency.strip().upper()
    totals: dict[str, float] = defaultdict(float)
    other: Counter = Counter()
    for cat, amount, c in rows:
        if not cat or not isinstance(amount, (int, float)):
            continue
        c = (str(c).strip().upper() if c else "") or cur
        if c == cur:
            totals[str(cat).strip()] += float(amount)
        else:
            other[c] += 1
    ordered = sorted(((k, v) for k, v in totals.items() if v > 0), key=lambda kv: (-kv[1], kv[0]))
    return ordered, dict(other)


def build_dashboard(
    wb: Workbook, data_sheet: str, rows: Iterable[tuple], currency: str
) -> int:
    """(Re)create the Dashboard sheet. Returns the number of categories charted."""
    cur = currency.strip().upper()
    if DASHBOARD in wb.sheetnames:
        idx = wb.sheetnames.index(DASHBOARD)
        del wb[DASHBOARD]
    else:
        idx = len(wb.sheetnames)
    ws = wb.create_sheet(DASHBOARD, idx)

    ordered, other = category_totals(rows, cur)

    ws["A1"] = "Spending Dashboard"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = f"All amounts in {cur}. Totals are formulas on the '{data_sheet}' sheet."
    ws["A2"].font = Font(italic=True, color="595959")
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 3

    for col, text in enumerate(["Category", f"Total ({cur})", "Share"], start=1):
        c = ws.cell(4, col, text)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = _HEADER_FILL
        c.alignment = Alignment(horizontal="center")

    if not ordered:
        ws["A5"] = f"No expenses in {cur} yet."
        return 0

    src = _quote(data_sheet)
    last = _FIRST_ROW + len(ordered) - 1
    total_row = last + 1
    for i, (cat, _) in enumerate(ordered):
        r = _FIRST_ROW + i
        ws.cell(r, 1, cat)
        ws.cell(
            r, 2,
            f'=SUMIFS({src}!$C:$C,{src}!$B:$B,$A{r},{src}!$D:$D,"{cur}")',
        ).number_format = "#,##0.00"
        ws.cell(r, 3, f"=IF($B${total_row}=0,0,B{r}/$B${total_row})").number_format = "0.0%"

    ws.cell(total_row, 1, "Total")
    ws.cell(total_row, 2, f"=SUM(B{_FIRST_ROW}:B{last})").number_format = "#,##0.00"
    ws.cell(total_row, 3, f"=SUM(C{_FIRST_ROW}:C{last})").number_format = "0.0%"
    for c in range(1, 4):
        cell = ws.cell(total_row, c)
        cell.font = Font(bold=True)
        cell.fill = _TOTAL_FILL
        cell.border = _TOP

    if other:
        listing = ", ".join(f"{k} ({v} rows)" for k, v in sorted(other.items()))
        ws.cell(total_row + 2, 1, f"Not included, other currencies: {listing}").font = Font(
            italic=True, color="595959"
        )

    chart = BarChart()
    chart.type = "bar"  # horizontal bars: category names stay readable
    chart.style = 10
    chart.title = CHART_TITLE
    chart.legend = None
    chart.add_data(Reference(ws, min_col=2, min_row=4, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=_FIRST_ROW, max_row=last))
    chart.x_axis.scaling.orientation = "maxMin"  # largest category on top
    chart.y_axis.crosses = "max"  # keep the value axis at the bottom after the flip
    chart.x_axis.delete = False  # openpyxl 3.1 hides axes unless told otherwise
    chart.y_axis.delete = False
    chart.y_axis.title = f"Amount ({cur})"
    chart.y_axis.number_format = "#,##0"
    chart.y_axis.majorGridlines = None
    chart.series[0].graphicalProperties.solidFill = "1F4E78"
    chart.gapWidth = 60
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.showSerName = False
    chart.dataLabels.showCatName = False
    chart.dataLabels.showLegendKey = False
    chart.dataLabels.showPercent = False
    chart.width = 18
    chart.height = min(14, max(7.5, 1.1 * len(ordered) + 3.5))
    ws.add_chart(chart, "E4")
    return len(ordered)
