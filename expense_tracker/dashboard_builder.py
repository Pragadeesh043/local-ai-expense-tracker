"""Dashboard sheet: 3 KPI cards, then a category table and a pie chart for each month.

The sheet is generated. It is deleted and rebuilt at the front of the workbook on every write,
so do not keep your own work on it. Every number on it is a formula on the other sheets.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import date
from typing import Iterable

from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook import Workbook
from openpyxl.worksheet.properties import PageSetupProperties

from .dates import cell_date
from .schema import Category, normalize_category

DASHBOARD = "Dashboard"
CATEGORIES = [c.value for c in Category]

# One fixed bright color per category, so FOOD is the same color in every month's chart.
CATEGORY_COLORS = {
    "FOOD": "FF595E",
    "GROCERY": "8AC926",
    "TRANSPORT": "1982C4",
    "PARKING": "6A4C93",
    "ENTERTAINMENT": "F15BB5",
    "PERSONAL": "FFCA3A",
    "OTHERS": "8D99AE",
}

_BLUE = "1F4E78"
_HEADER_FILL = PatternFill("solid", fgColor=_BLUE)
_TOTAL_FILL = PatternFill("solid", fgColor="EDEDED")
_CARD_FILL = PatternFill("solid", fgColor="F7F9FC")
_BLOCK_ROWS = 17  # rows reserved per month: heading, table, and room for the chart
_FIRST_MONTH_ROW = 8


@dataclass
class DashboardInfo:
    months: int = 0
    charts: int = 0


def _quote(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"


def month_tables(rows: Iterable[tuple], currency: str) -> tuple[dict[str, dict[str, float]], dict[str, int]]:
    """Group rows by month. Returns ({"2026-11": {category: total}}, other-currency row counts).

    rows: (date, category, amount, currency). Months are newest first. Every month lists all 7
    categories, with 0 for the ones that had no spending.
    """
    cur = currency.strip().upper()
    months: dict[str, dict[str, float]] = {}
    other: Counter = Counter()
    for d, cat, amount, c in rows:
        day = cell_date(d)
        if day is None or not isinstance(amount, (int, float)):
            continue
        cats = months.setdefault(day.strftime("%Y-%m"), {name: 0.0 for name in CATEGORIES})
        cc = (str(c).strip().upper() if c else "") or cur
        if cc != cur:
            other[cc] += 1
            continue
        cats[normalize_category(cat).value] += float(amount)
    return dict(sorted(months.items(), reverse=True)), dict(other)


def _card(ws, first_col: int, label: str, formula: str, color: str) -> None:
    """A KPI card: a label row and a large value row, each merged across 3 columns, with a box border."""
    side = Side(style="medium", color=_BLUE)
    last_col = first_col + 2
    top = ws.cell(4, first_col, label)
    top.font = Font(bold=True, size=10, color="595959")
    top.alignment = Alignment(horizontal="center", vertical="center")
    top.fill = _CARD_FILL
    top.border = Border(left=side, right=side, top=side)
    val = ws.cell(5, first_col, formula)
    val.font = Font(bold=True, size=22, color=color)
    val.alignment = Alignment(horizontal="center", vertical="center")
    val.number_format = "#,##0.00"
    val.fill = _CARD_FILL
    val.border = Border(left=side, right=side, bottom=side)
    # Borders are set on the top-left cell first. merge_cells copies them to the edge cells.
    ws.merge_cells(start_row=4, start_column=first_col, end_row=4, end_column=last_col)
    ws.merge_cells(start_row=5, start_column=first_col, end_row=5, end_column=last_col)


def build_dashboard(
    wb: Workbook, data_sheet: str, income_sheet: str, rows: Iterable[tuple], currency: str
) -> DashboardInfo:
    """(Re)create the Dashboard as the first sheet. rows: (date, category, amount, currency)."""
    cur = currency.strip().upper()
    if DASHBOARD in wb.sheetnames:
        del wb[DASHBOARD]
    ws = wb.create_sheet(DASHBOARD, 0)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = _BLUE
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)  # print one page wide
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    for col, width in {"A": 16, "B": 14, "C": 10, "D": 3, "H": 3}.items():
        ws.column_dimensions[col].width = width
    for col in "EFGIJK":
        ws.column_dimensions[col].width = 13

    src, inc = _quote(data_sheet), _quote(income_sheet)

    ws["A1"] = "Financial Dashboard"
    ws["A1"].font = Font(bold=True, size=18, color=_BLUE)
    ws["A2"] = f"Amounts in {cur}. Every figure is a live formula. Charts are rebuilt each time you add an entry."
    ws["A2"].font = Font(italic=True, color="595959")

    # A. KPI cards. Total Spent skips the monthly TOTAL rows (they have no category), so nothing is counted twice.
    _card(ws, 1, f"TOTAL INCOME ({cur})", f"=SUM({inc}!$C:$C)", "2E7D32")
    _card(
        ws, 5, f"TOTAL SPENT ({cur})",
        f'=SUMIFS({src}!$C:$C,{src}!$B:$B,"<>",{src}!$D:$D,"{cur}")', "C0392B",
    )
    _card(ws, 9, f"NET BALANCE ({cur})", "=A5-E5", "2E7D32")
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 40
    ws.conditional_formatting.add(
        "I5", CellIsRule(operator="lessThan", formula=["0"], font=Font(bold=True, size=22, color="C0392B"))
    )

    # B. One block per month.
    months, other = month_tables(rows, cur)
    if other:
        listing = ", ".join(f"{k} ({v} rows)" for k, v in sorted(other.items()))
        ws["A6"] = f"Not included, other currencies: {listing}"
        ws["A6"].font = Font(italic=True, color="595959")
    if not months:
        ws.cell(_FIRST_MONTH_ROW, 1, "No expenses yet.")
        return DashboardInfo()

    charts = 0
    r0 = _FIRST_MONTH_ROW
    for key, totals in months.items():
        y, m = (int(x) for x in key.split("-"))
        ny, nm = (y + 1, 1) if m == 12 else (y, m + 1)
        label = date(y, m, 1).strftime("%B %Y")
        ranked = sorted(totals.items(), key=lambda kv: (-kv[1], CATEGORIES.index(kv[0])))
        spent = [(c, t) for c, t in ranked if t > 0]

        heading = ws.cell(r0, 1, label)
        heading.font = Font(bold=True, size=14, color=_BLUE)
        for c in range(1, 12):
            ws.cell(r0, c).border = Border(bottom=Side(style="thin", color=_BLUE))

        hr, first = r0 + 1, r0 + 2
        total_row = first + len(ranked)
        for col, text in enumerate(["Category", f"Spent ({cur})", "Share"], start=1):
            h = ws.cell(hr, col, text)
            h.font = Font(bold=True, color="FFFFFF")
            h.fill = _HEADER_FILL
            h.alignment = Alignment(horizontal="center")
        for i, (cat, _) in enumerate(ranked):
            r = first + i
            ws.cell(r, 1, cat)
            ws.cell(
                r, 2,
                f'=SUMIFS({src}!$C:$C,{src}!$B:$B,$A{r},'
                f'{src}!$A:$A,">="&DATE({y},{m},1),{src}!$A:$A,"<"&DATE({ny},{nm},1),'
                f'{src}!$D:$D,"{cur}")',
            ).number_format = "#,##0.00"
            ws.cell(r, 3, f"=IF($B${total_row}=0,0,B{r}/$B${total_row})").number_format = "0.0%"
            ws.cell(r, 1).font = Font(color="000000", bold=False)
        ws.cell(total_row, 1, "Total")
        ws.cell(total_row, 2, f"=SUM(B{first}:B{total_row - 1})").number_format = "#,##0.00"
        ws.cell(total_row, 3, f"=SUM(C{first}:C{total_row - 1})").number_format = "0.0%"
        for c in range(1, 4):
            cell = ws.cell(total_row, c)
            cell.font = Font(bold=True)
            cell.fill = _TOTAL_FILL
            cell.border = Border(top=Side(style="thin"))

        if spent:
            k = len(spent)  # rows are sorted by spend, so the spending categories come first
            pie = PieChart()
            pie.title = f"{label}: Spending by Category"
            pie.add_data(Reference(ws, min_col=2, min_row=hr, max_row=first + k - 1), titles_from_data=True)
            pie.set_categories(Reference(ws, min_col=1, min_row=first, max_row=first + k - 1))
            for i, (cat, _) in enumerate(spent):
                pt = DataPoint(idx=i)
                pt.graphicalProperties.solidFill = CATEGORY_COLORS[cat]
                pt.graphicalProperties.line.solidFill = "FFFFFF"
                pie.series[0].dPt.append(pt)
            pie.dataLabels = DataLabelList()
            pie.dataLabels.showPercent = True
            pie.dataLabels.showVal = False
            pie.dataLabels.showCatName = False
            pie.dataLabels.showSerName = False
            pie.dataLabels.showLegendKey = False
            pie.dataLabels.dLblPos = "outEnd"
            pie.legend.position = "r"
            pie.width = 15
            pie.height = 7.6
            ws.add_chart(pie, f"E{hr}")
            charts += 1
        else:
            ws.cell(hr + 1, 5, f"No spending in {cur} this month.").font = Font(italic=True, color="595959")

        r0 += _BLOCK_ROWS
    return DashboardInfo(months=len(months), charts=charts)
