"""Write expenses to an .xlsx file, kept in date order. Creates it with styled headers if missing."""

from __future__ import annotations

from copy import copy
from datetime import date as date_cls
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .schema import Expense

HEADERS = ["Date", "Category", "Amount", "Currency", "Merchant", "Payment Method", "Raw Note"]
_WIDTHS = [12, 18, 12, 10, 20, 16, 50]
DATE_COL = 1
AMOUNT_COL = 3


class SheetError(RuntimeError):
    pass


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


def _sort_key(v, order: str) -> tuple[int, int]:
    """Sort key for one row's date cell. Rows without a readable date always go last."""
    d = _date_of(v)
    if d is None:
        return (1, 0)
    return (0, -d.toordinal() if order == "desc" else d.toordinal())


def _norm_order(order: str) -> str:
    return "asc" if str(order).strip().lower().startswith("asc") else "desc"


def append_expenses(
    path: Path, sheet_name: str, expenses: list[Expense], order: str = "desc"
) -> list[int]:
    """Add expenses, then re-sort every data row by date. One save for everything.

    order: "desc" (newest on top, default) or "asc" (oldest first).
    Rows with the same date keep their existing order, and new rows go after them.
    Header, freeze pane and column widths are not touched. Each row keeps its own cell
    styles when it moves. Formulas in moved rows are re-pointed to their new row.

    Returns the final row number of each new expense, in the order given.
    If the file is locked, nothing is written.
    """
    if not expenses:
        return []
    order = _norm_order(order)
    wb, ws = _open(path, sheet_name)
    ncols = max(len(HEADERS), ws.max_column)
    old_max = ws.max_row

    # 1. Collect existing data rows (values, styles, origin row). Blank rows are dropped.
    entries: list[dict] = []
    for r in range(2, old_max + 1):
        cells = [ws.cell(r, c) for c in range(1, ncols + 1)]
        if all(c.value is None for c in cells):
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

    # 4. Write back from row 2. Styles travel with their row.
    new_rows = [0] * len(expenses)
    for pos, en in enumerate(entries):
        r = pos + 2
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

    # 5. Remove leftover rows if blank rows were dropped.
    last = len(entries) + 1
    if old_max > last:
        ws.delete_rows(last + 1, old_max - last)

    try:
        wb.save(path)
    except PermissionError as err:
        raise SheetError(f"Cannot write {path}. Close it in Excel and retry.") from err
    return new_rows


def append_expense(path: Path, sheet_name: str, e: Expense, order: str = "desc") -> int:
    """Add one expense. Returns its final row number."""
    return append_expenses(path, sheet_name, [e], order)[0]


def read_recent(path: Path, sheet_name: str, n: int = 10) -> list[list]:
    """The n most recent rows by date, oldest first. Empty list if the file or sheet is missing."""
    if not path.exists():
        return []
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except PermissionError as e:
        raise SheetError(f"Cannot open {path}. Close it in Excel first.") from e
    if sheet_name not in wb.sheetnames:
        return []
    rows = [list(r) for r in wb[sheet_name].iter_rows(min_row=2, values_only=True) if any(r)]
    wb.close()
    rows.sort(key=lambda r: _date_of(r[0]) or date_cls.min)  # stable: same-day rows keep sheet order
    return rows[-n:]
