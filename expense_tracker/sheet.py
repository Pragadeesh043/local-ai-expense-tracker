"""Write expenses to an .xlsx file, kept in date order, with optional monthly totals and a Dashboard.

Monthly TOTAL rows and the Dashboard sheet are generated output. Every write strips the old
TOTAL rows, sorts the real expense rows, and rebuilds the TOTAL rows and Dashboard in the same
save, so formula ranges never go stale.
"""

from __future__ import annotations

import re
from copy import copy
from dataclasses import dataclass, field
from datetime import date as date_cls
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .chart_builder import DASHBOARD, build_dashboard
from .schema import Expense

HEADERS = ["Date", "Category", "Amount", "Currency", "Merchant", "Payment Method", "Raw Note"]
_WIDTHS = [12, 18, 12, 10, 20, 16, 50]
DATE_COL = 1
CATEGORY_COL = 2
AMOUNT_COL = 3
CURRENCY_COL = 4

_SUMMARY_RE = re.compile(r"^TOTAL \((\d{4}-\d{2})\)$")
_TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
_TOTAL_FONT = Font(bold=True)
_TOTAL_BORDER = Border(top=Side(style="thin"), bottom=Side(style="thin"))


class SheetError(RuntimeError):
    pass


@dataclass
class WriteResult:
    rows: list[int] = field(default_factory=list)  # final row of each new expense, in input order
    months: int = 0  # distinct months that have TOTAL rows after this write
    categories: int = 0  # bars on the Dashboard chart (0 if no Dashboard was written)
    totals_on: bool = False
    dashboard_written: bool = False


def _style_header(ws: Worksheet) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    for i, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = _WIDTHS[i - 1]
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
            _style_header(ws)
        if ws.max_row == 1 and ws.cell(1, 1).value is None:
            _style_header(ws)
        return wb, ws

    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    _style_header(ws)
    return wb, ws


def _date_of(v) -> date_cls | None:
    """Best-effort date from a cell value. None if it is not a date."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date_cls):
        return v
    if isinstance(v, str):
        try:
            return date_cls.fromisoformat(v.strip()[:10])
        except ValueError:
            return None
    return None


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


def _apply(
    path: Path,
    sheet_name: str,
    expenses: list[Expense],
    order: str,
    currency: str,
    totals: bool | None,
    dashboard: bool | None,
) -> WriteResult:
    """The single write path.

    totals:    True = write monthly TOTAL rows, False = remove them, None = keep as they are.
    dashboard: True = (re)build the Dashboard, False = skip, None = refresh only if it exists.
    Everything is written in one save. If the file is locked, nothing is written.
    """
    order = _norm_order(order)
    currency = currency.strip().upper() or "INR"
    wb, ws = _open(path, sheet_name)
    ncols = max(len(HEADERS), ws.max_column)
    old_max = ws.max_row

    # 1. Collect existing data rows. Blank rows and old TOTAL rows are dropped.
    entries: list[dict] = []
    had_totals = False
    for r in range(2, old_max + 1):
        cells = [ws.cell(r, c) for c in range(1, ncols + 1)]
        if all(c.value is None for c in cells):
            continue
        if _is_summary(cells[DATE_COL - 1].value):
            had_totals = True
            continue
        entries.append(
            {
                "vals": [c.value for c in cells],
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
    with_totals = had_totals if totals is None else totals
    layout = _layout(entries, with_totals, currency)
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

    # 6. Dashboard.
    result = WriteResult(
        rows=new_rows,
        months=len({it["month"] for it in layout if it["kind"] == "total"}),
        totals_on=with_totals,
    )
    if dashboard is True or (dashboard is None and DASHBOARD in wb.sheetnames):
        rows = [(en["vals"][CATEGORY_COL - 1], en["vals"][AMOUNT_COL - 1], en["vals"][CURRENCY_COL - 1])
                for en in entries]
        result.categories = build_dashboard(wb, sheet_name, rows, currency)
        result.dashboard_written = True

    # Formulas have no stored results when written by openpyxl. Ask Excel to calculate on open.
    if wb.calculation is not None:
        wb.calculation.fullCalcOnLoad = True

    try:
        wb.save(path)
    except PermissionError as err:
        raise SheetError(f"Cannot write {path}. Close it in Excel and retry.") from err
    return result


def append_expenses(
    path: Path,
    sheet_name: str,
    expenses: list[Expense],
    order: str = "desc",
    currency: str = "INR",
) -> list[int]:
    """Add expenses, then re-sort every data row by date. One save for everything.

    order: "desc" (newest on top, default) or "asc" (oldest first).
    Rows with the same date keep their existing order, and new rows go after them.
    Header, freeze pane and column widths are not touched. Each row keeps its own cell
    styles when it moves. Formulas in moved rows are re-pointed to their new row.
    If the sheet already has monthly TOTAL rows or a Dashboard, they are rebuilt.

    Returns the final row number of each new expense, in the order given.
    """
    if not expenses:
        return []
    return _apply(path, sheet_name, expenses, order, currency, None, None).rows


def append_expense(
    path: Path, sheet_name: str, e: Expense, order: str = "desc", currency: str = "INR"
) -> int:
    """Add one expense. Returns its final row number."""
    return append_expenses(path, sheet_name, [e], order, currency)[0]


def summarize(
    path: Path,
    sheet_name: str,
    order: str = "desc",
    currency: str = "INR",
    totals: bool = True,
    dashboard: bool = True,
) -> WriteResult:
    """Write monthly TOTAL rows and the Dashboard for an existing file."""
    if not path.exists():
        raise SheetError(f"{path} does not exist. Add an expense first.")
    return _apply(path, sheet_name, [], order, currency, totals, dashboard)


def remove_totals(path: Path, sheet_name: str, order: str = "desc", currency: str = "INR") -> WriteResult:
    """Remove the monthly TOTAL rows. The Dashboard is refreshed if it exists."""
    if not path.exists():
        raise SheetError(f"{path} does not exist.")
    return _apply(path, sheet_name, [], order, currency, False, None)


def read_recent(path: Path, sheet_name: str, n: int = 10) -> list[list]:
    """The n most recent expense rows by date, oldest first. TOTAL rows are skipped."""
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
