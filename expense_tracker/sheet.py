"""Write expenses and income to an .xlsx file and keep the generated parts in step.

Sheets: Dashboard (first, generated), Expenses, Income. Every write does everything in one save:
  1. Existing expense rows are collected. Old monthly TOTAL rows are dropped.
  2. New rows are added, then all rows are sorted by date (stable).
  3. Monthly TOTAL rows are rebuilt (unless EXPENSE_TOTALS=0).
  4. The Income sheet is created if missing, then sorted by date.
  5. The Dashboard is rebuilt.
If the file is locked, nothing is written.
"""

from __future__ import annotations

import re
from copy import copy
from dataclasses import dataclass, field
from datetime import date as date_cls
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .dashboard_builder import DASHBOARD, build_dashboard
from .dates import cell_date as _date_of
from .income import Income
from .schema import Expense, normalize_category

HEADERS = ["Date", "Category", "Amount", "Currency", "Merchant", "Payment Method", "Raw Note"]
_WIDTHS = [12, 18, 12, 10, 20, 16, 50]
INCOME = "Income"
INCOME_HEADERS = ["Date (YYYY-MM-DD)", "Source", "Amount (INR)", "Notes"]
_INCOME_WIDTHS = [18, 24, 16, 40]
DATE_COL = 1
CATEGORY_COL = 2
AMOUNT_COL = 3
CURRENCY_COL = 4

_SUMMARY_RE = re.compile(r"^TOTAL \((\d{4}-\d{2})\)$")
_TOTAL_FILL = PatternFill("solid", fgColor="EDEDED")
_TOTAL_FONT = Font(bold=True)
_TOTAL_BORDER = Border(top=Side(style="thin"), bottom=Side(style="thin"))


class SheetError(RuntimeError):
    pass


@dataclass
class WriteResult:
    rows: list[int] = field(default_factory=list)  # final row of each new expense, in input order
    income_rows: list[int] = field(default_factory=list)  # final row of each new income, in input order
    months: int = 0  # months shown on the Dashboard
    charts: int = 0  # pie charts on the Dashboard
    totals_on: bool = False


def _style_header(ws: Worksheet, headers: list[str], widths: list[int], color: str = "1F4E78") -> None:
    fill = PatternFill("solid", fgColor=color)
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.freeze_panes = "A2"


def _open(path: Path, sheet_name: str):
    if path.exists():
        try:
            wb = load_workbook(path)
        except PermissionError as e:
            raise SheetError(f"Cannot open {path}. Close it in Excel first.") from e
        except Exception as e:  # corrupt or not an xlsx
            raise SheetError(f"{path} is not a readable .xlsx file: {e}") from e
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(sheet_name)
            _style_header(ws, HEADERS, _WIDTHS)
        if ws.max_row == 1 and ws.cell(1, 1).value is None:
            _style_header(ws, HEADERS, _WIDTHS)
        return wb, ws

    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    _style_header(ws, HEADERS, _WIDTHS)
    return wb, ws


def _ensure_income(wb: Workbook, after: Worksheet) -> Worksheet:
    """The Income sheet, created right after the Expenses sheet if it does not exist."""
    if INCOME in wb.sheetnames:
        ws = wb[INCOME]
    else:
        ws = wb.create_sheet(INCOME, wb.sheetnames.index(after.title) + 1)
    if ws.max_row == 1 and ws.cell(1, 1).value is None:
        _style_header(ws, INCOME_HEADERS, _INCOME_WIDTHS, color="2E7D32")
    return ws


def _month_of(v) -> str | None:
    d = _date_of(v)
    return d.strftime("%Y-%m") if d else None


def _is_summary(v) -> bool:
    return isinstance(v, str) and bool(_SUMMARY_RE.match(v.strip()))


def _sort_key(v, order: str) -> tuple[int, int]:
    """Sort key for one row's date cell. Rows without a readable date always go last."""
    d = _date_of(v)
    if d is None:
        return (1, 0)
    return (0, -d.toordinal() if order == "desc" else d.toordinal())


def _norm_order(order: str) -> str:
    return "asc" if str(order).strip().lower().startswith("asc") else "desc"


def _currency_of(entry: dict, default: str) -> str:
    v = entry["vals"][CURRENCY_COL - 1]
    return (str(v).strip().upper() if v else "") or default


def _migrate_category(v):
    """Old labels such as 'Food & Dining' become one of the 7 categories. Blank cells stay blank."""
    if isinstance(v, str) and v.strip():
        return normalize_category(v).value
    return v


def _layout(entries: list[dict], with_totals: bool, default_currency: str) -> list[dict]:
    """Final top-to-bottom row plan: data rows, with TOTAL rows after each month's block.

    A month with several currencies gets one TOTAL row per currency, because adding INR to USD
    is meaningless.
    """
    if not with_totals:
        return [{"kind": "data", "en": en} for en in entries]

    out: list[dict] = []
    i, n = 0, len(entries)
    while i < n:
        month = _month_of(entries[i]["vals"][DATE_COL - 1])
        j = i + 1
        if month is not None:
            while j < n and _month_of(entries[j]["vals"][DATE_COL - 1]) == month:
                j += 1
        first_idx = len(out)
        out.extend({"kind": "data", "en": en} for en in entries[i:j])
        if month is not None:
            currencies = sorted({_currency_of(en, default_currency) for en in entries[i:j]})
            for cur in currencies:
                out.append(
                    {
                        "kind": "total",
                        "month": month,
                        "currency": cur,
                        "multi": len(currencies) > 1,
                        "first_idx": first_idx,
                        "last_idx": first_idx + (j - i) - 1,
                    }
                )
        i = j
    return out


def _write_total(ws: Worksheet, r: int, ncols: int, item: dict) -> None:
    first, last = item["first_idx"] + 2, item["last_idx"] + 2  # list index -> sheet row
    for c in range(1, ncols + 1):
        cell = ws.cell(r, c)
        cell.value = None
        cell.style = "Normal"
    ws.cell(r, DATE_COL).value = f"TOTAL ({item['month']})"
    if item["multi"]:
        ws.cell(r, AMOUNT_COL).value = (
            f'=SUMIF(D{first}:D{last},"{item["currency"]}",C{first}:C{last})'
        )
    else:
        ws.cell(r, AMOUNT_COL).value = f"=SUM(C{first}:C{last})"
    ws.cell(r, CURRENCY_COL).value = item["currency"]
    for c in range(1, ncols + 1):
        cell = ws.cell(r, c)
        cell.font = _TOTAL_FONT
        cell.fill = _TOTAL_FILL
        cell.border = _TOTAL_BORDER
    ws.cell(r, AMOUNT_COL).number_format = "#,##0.00"


def _rewrite_income(ws: Worksheet, incomes: list[Income], order: str) -> list[int]:
    """Add income rows and re-sort the sheet by date. Returns each new row's final row number."""
    ncols = max(len(INCOME_HEADERS), ws.max_column)
    old_max = ws.max_row
    entries: list[dict] = []
    for r in range(2, old_max + 1):
        cells = [ws.cell(r, c) for c in range(1, ncols + 1)]
        if all(c.value is None for c in cells):
            continue
        entries.append(
            {"vals": [c.value for c in cells], "styles": [copy(c._style) for c in cells], "new": None}
        )
    for i, inc in enumerate(incomes):
        vals = [date_cls.fromisoformat(inc.date), inc.source, inc.amount, inc.notes or None]
        vals += [None] * (ncols - len(vals))
        entries.append({"vals": vals, "styles": None, "new": i})

    entries.sort(key=lambda en: _sort_key(en["vals"][0], order))

    new_rows = [0] * len(incomes)
    for pos, en in enumerate(entries):
        r = pos + 2
        for c, v in enumerate(en["vals"], start=1):
            cell = ws.cell(r, c)
            cell.value = v
            if en["styles"] is not None:
                cell._style = copy(en["styles"][c - 1])
            else:
                cell.style = "Normal"
        if en["new"] is not None:
            ws.cell(r, 1).number_format = "YYYY-MM-DD"
            ws.cell(r, 3).number_format = "#,##0.00"
            new_rows[en["new"]] = r
    last = len(entries) + 1
    if old_max > last:
        ws.delete_rows(last + 1, old_max - last)
    return new_rows


def _apply(
    path: Path,
    sheet_name: str,
    expenses: list[Expense],
    incomes: list[Income],
    order: str,
    currency: str,
    totals: bool,
) -> WriteResult:
    order = _norm_order(order)
    currency = currency.strip().upper() or "INR"
    wb, ws = _open(path, sheet_name)
    income_ws = _ensure_income(wb, ws)
    ncols = max(len(HEADERS), ws.max_column)
    old_max = ws.max_row

    # 1. Collect existing expense rows. Blank rows and old TOTAL rows are dropped.
    entries: list[dict] = []
    for r in range(2, old_max + 1):
        cells = [ws.cell(r, c) for c in range(1, ncols + 1)]
        if all(c.value is None for c in cells):
            continue
        if _is_summary(cells[DATE_COL - 1].value):
            continue
        vals = [c.value for c in cells]
        vals[CATEGORY_COL - 1] = _migrate_category(vals[CATEGORY_COL - 1])
        entries.append(
            {
                "vals": vals,
                "styles": [copy(c._style) for c in cells],
                "origin_row": r,
                "new": None,
            }
        )

    # 2. Add the new rows.
    for i, e in enumerate(expenses):
        vals = [
            date_cls.fromisoformat(e.date),
            e.category,
            e.amount,
            e.currency,
            e.merchant,
            e.payment_method,
            e.raw_note,
        ]
        vals += [None] * (ncols - len(vals))
        entries.append({"vals": vals, "styles": None, "origin_row": None, "new": i})

    # 3. Stable sort by date. Python's sort keeps the order of equal keys.
    entries.sort(key=lambda en: _sort_key(en["vals"][DATE_COL - 1], order))

    # 4. Plan the final rows, then write them from row 2. Styles travel with their row.
    layout = _layout(entries, totals, currency)
    new_rows = [0] * len(expenses)
    for k, item in enumerate(layout):
        r = k + 2
        if item["kind"] == "total":
            _write_total(ws, r, ncols, item)
            continue
        en = item["en"]
        for c, v in enumerate(en["vals"], start=1):
            cell = ws.cell(r, c)
            if isinstance(v, str) and v.startswith("=") and en["origin_row"] not in (None, r):
                col = get_column_letter(c)
                v = Translator(v, origin=f"{col}{en['origin_row']}").translate_formula(f"{col}{r}")
            cell.value = v
            if en["styles"] is not None:
                cell._style = copy(en["styles"][c - 1])
            else:
                cell.style = "Normal"
        if en["new"] is not None:
            ws.cell(r, DATE_COL).number_format = "YYYY-MM-DD"
            ws.cell(r, AMOUNT_COL).number_format = "#,##0.00"
            new_rows[en["new"]] = r

    # 5. Remove leftover rows (blank rows or removed TOTAL rows).
    last = len(layout) + 1
    if old_max > last:
        ws.delete_rows(last + 1, old_max - last)

    # 6. Income sheet, then the Dashboard (always the first sheet).
    income_rows = _rewrite_income(income_ws, incomes, order)
    dash_rows = [
        (en["vals"][DATE_COL - 1], en["vals"][CATEGORY_COL - 1], en["vals"][AMOUNT_COL - 1], en["vals"][CURRENCY_COL - 1])
        for en in entries
    ]
    info = build_dashboard(wb, sheet_name, INCOME, dash_rows, currency)

    # Open on the Dashboard, with only that tab selected (several selected tabs would group the sheets).
    wb.active = wb.sheetnames.index(DASHBOARD)
    for sheet in wb.worksheets:
        sheet.sheet_view.tabSelected = sheet.title == DASHBOARD

    # Formulas have no stored results when written by openpyxl. Ask Excel to calculate on open.
    if wb.calculation is not None:
        wb.calculation.fullCalcOnLoad = True

    try:
        wb.save(path)
    except PermissionError as err:
        raise SheetError(f"Cannot write {path}. Close it in Excel and retry.") from err
    return WriteResult(
        rows=new_rows, income_rows=income_rows, months=info.months, charts=info.charts, totals_on=totals
    )


def append_expenses(
    path: Path,
    sheet_name: str,
    expenses: list[Expense],
    order: str = "desc",
    currency: str = "INR",
    totals: bool = True,
) -> list[int]:
    """Add expenses, re-sort by date, rebuild monthly totals, the Income sheet and the Dashboard.

    order: "desc" (newest on top, default) or "asc" (oldest first).
    Rows with the same date keep their existing order, and new rows go after them.
    Returns the final row number of each new expense, in the order given.
    """
    if not expenses:
        return []
    return _apply(path, sheet_name, expenses, [], order, currency, totals).rows


def append_expense(
    path: Path, sheet_name: str, e: Expense, order: str = "desc", currency: str = "INR", totals: bool = True
) -> int:
    """Add one expense. Returns its final row number."""
    return append_expenses(path, sheet_name, [e], order, currency, totals)[0]


def append_incomes(
    path: Path,
    sheet_name: str,
    incomes: list[Income],
    order: str = "desc",
    currency: str = "INR",
    totals: bool = True,
) -> list[int]:
    """Add income rows to the Income sheet (created if missing) and rebuild the Dashboard.

    Returns the final row number of each new income, in the order given.
    """
    if not incomes:
        return []
    return _apply(path, sheet_name, [], incomes, order, currency, totals).income_rows


def summarize(
    path: Path, sheet_name: str, order: str = "desc", currency: str = "INR", totals: bool = True
) -> WriteResult:
    """Rebuild monthly totals, the Income sheet and the Dashboard for an existing file."""
    if not path.exists():
        raise SheetError(f"{path} does not exist. Add an expense first.")
    return _apply(path, sheet_name, [], [], order, currency, totals)


def read_recent(path: Path, sheet_name: str, n: int = 10) -> list[list]:
    """The n most recent rows of a sheet by date, oldest first. TOTAL rows are skipped."""
    if not path.exists():
        return []
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except PermissionError as e:
        raise SheetError(f"Cannot open {path}. Close it in Excel first.") from e
    if sheet_name not in wb.sheetnames:
        return []
    rows = [
        list(r)
        for r in wb[sheet_name].iter_rows(min_row=2, values_only=True)
        if any(r) and not _is_summary(r[0])
    ]
    wb.close()
    rows.sort(key=lambda r: _date_of(r[0]) or date_cls.min)  # stable: same-day rows keep sheet order
    return rows[-n:]
