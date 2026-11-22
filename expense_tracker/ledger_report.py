"""Shared reporting and atomic publication of the generated Excel workbook."""
from __future__ import annotations

import os
from collections import Counter
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

from .ledger import Entry, Ledger, LedgerError


_MARKER = "_Tracker"
_REPORT_VERSION = "2"
_BLUE = "17344C"
_MONEY = '#,##0.00;[Red](#,##0.00);0.00'


def totals(entries: list[Entry]) -> tuple[int, int]:
    return (
        sum(item.cents for item in entries if item.kind == "income"),
        sum(item.cents for item in entries if item.kind == "expense"),
    )


@dataclass(frozen=True)
class MonthlyBalance:
    opening: int
    income: int
    spending: int

    @property
    def net_change(self) -> int:
        return self.income - self.spending

    @property
    def closing(self) -> int:
        return self.opening + self.net_change


def month_groups(entries: list[Entry]) -> dict[tuple[str, str], list[Entry]]:
    groups: dict[tuple[str, str], list[Entry]] = {}
    for item in entries:
        groups.setdefault((item.day[:7], item.currency), []).append(item)
    return groups


def monthly_balances(groups: dict[tuple[str, str], list[Entry]]) -> dict[tuple[str, str], MonthlyBalance]:
    result: dict[tuple[str, str], MonthlyBalance] = {}
    closing_by_currency: dict[str, int] = {}
    for key, rows in sorted(groups.items()):
        currency = key[1]
        income, spending = totals(rows)
        balance = MonthlyBalance(closing_by_currency.get(currency, 0), income, spending)
        result[key] = balance
        closing_by_currency[currency] = balance.closing
    return result


def breakdown(entries: list[Entry], field: str) -> list[tuple[str, int]]:
    buckets: Counter[str] = Counter()
    for item in entries:
        if item.kind == "expense":
            buckets[getattr(item, field)] += item.cents
    return sorted(buckets.items(), key=lambda pair: (-pair[1], pair[0].casefold()))


def _append(sheet, values) -> None:
    sheet.append([ILLEGAL_CHARACTERS_RE.sub("", value) if isinstance(value, str) else value for value in values])
    for cell in sheet[sheet.max_row]:
        if isinstance(cell.value, str):
            cell.data_type = "s"


def _table(sheet, name: str, headers: tuple[str, ...], widths: tuple[int, ...]) -> None:
    for column, title, width in zip(sheet[1], headers, widths, strict=True):
        column.value = title
        column.fill = PatternFill("solid", fgColor=_BLUE)
        column.font = Font(color="FFFFFF", bold=True)
        sheet.column_dimensions[column.column_letter].width = width
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    if sheet.max_row > 1:
        table = Table(displayName=name, ref=sheet.dimensions)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        sheet.add_table(table)


def _transaction_sheet(workbook: Workbook, entries: list[Entry], kind: str) -> None:
    expense = kind == "expense"
    sheet = workbook.create_sheet("Expenses" if expense else "Income")
    headers = ("Date", "Category", "Amount", "Currency", "Merchant", "Payment Method", "Raw Note") if expense else (
        "Date", "Source", "Amount", "Currency", "Notes"
    )
    _append(sheet, headers)
    for item in entries:
        if item.kind != kind:
            continue
        day = date.fromisoformat(item.day)
        amount = item.cents / 100
        _append(sheet, (day, item.category, amount, item.currency, item.name, item.payment, item.note) if expense else (
            day, item.name, amount, item.currency, item.note
        ))
        sheet.cell(sheet.max_row, 1).number_format = "yyyy-mm-dd"
        sheet.cell(sheet.max_row, 3).number_format = _MONEY
    widths = (14, 19, 18, 12, 27, 20, 52) if expense else (14, 29, 18, 12, 52)
    _table(sheet, kind.title() + "Ledger", headers, widths)


def _dashboard(sheet, groups: dict[tuple[str, str], list[Entry]],
               balances: dict[tuple[str, str], MonthlyBalance], currency: str, month: str) -> None:
    entries = groups[(month, currency)]
    balance = balances[(month, currency)]
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A7"
    sheet.sheet_properties.tabColor = "167C73"
    sheet.merge_cells("A1:I1")
    title = sheet["A1"]
    title.value = "Financial Dashboard"
    title.font = Font(size=22, bold=True, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor=_BLUE)
    title.alignment = Alignment(indent=1, vertical="center")
    sheet.row_dimensions[1].height = 36
    sheet["A2"] = f"{date.fromisoformat(month + '-01'):%B %Y}  |  {currency}"
    sheet["A2"].font = Font(bold=True, color=_BLUE)
    for col, (label, value) in zip("ACEGI", (
        ("Opening balance", balance.opening / 100), ("Income", balance.income / 100),
        ("Spending", balance.spending / 100), ("Closing balance", balance.closing / 100),
        ("Savings rate", balance.net_change / balance.income if balance.income else None),
    ), strict=True):
        sheet[col + "4"] = label
        sheet[col + "4"].font = Font(bold=True, color=_BLUE)
        cell = sheet[col + "5"]
        cell.value = "—" if value is None else value
        cell.number_format = "0.0%" if col == "I" else _MONEY
        cell.font = Font(bold=True, size=16, color="167C73" if value is None or value >= 0 else "B33238")
        cell.fill = PatternFill("solid", fgColor="E9F5F3")
    sheet.row_dimensions[5].height = 26
    for col, width in zip("ABCDEFGHI", (22, 16, 24, 16, 24, 16, 19, 16, 16), strict=True):
        sheet.column_dimensions[col].width = width

    for start, field, label in ((1, "category", "Category"), (3, "name", "Merchant"), (5, "payment", "Payment method")):
        rows = breakdown(entries, field)[:10]
        for offset, heading in enumerate((label, f"Spent ({currency})")):
            cell = sheet.cell(8, start + offset, heading)
            cell.fill = PatternFill("solid", fgColor=_BLUE)
            cell.font = Font(bold=True, color="FFFFFF")
        for index, (name, cents) in enumerate(rows, 9):
            cell = sheet.cell(index, start, ILLEGAL_CHARACTERS_RE.sub("", name))
            cell.data_type = "s"
            sheet.cell(index, start + 1, cents / 100).number_format = _MONEY
        if field == "category" and rows:
            chart = BarChart()
            chart.type = "bar"
            chart.title = "Spending by category"
            chart.legend = None
            chart.add_data(Reference(sheet, min_col=2, min_row=8, max_row=8 + len(rows)), titles_from_data=True)
            chart.set_categories(Reference(sheet, min_col=1, min_row=9, max_row=8 + len(rows)))
            sheet.add_chart(chart, "A22")

    history = [(key[0], value) for key, value in balances.items() if key[1] == currency][-12:]
    for column, label in enumerate(("Month", "Income", "Spending", "Closing balance"), 1):
        sheet.cell(39, column, label).font = Font(bold=True, color=_BLUE)
    for row, (label, value) in enumerate(history, 40):
        sheet.cell(row, 1, label)
        for column, cents in enumerate((value.income, value.spending, value.closing), 2):
            sheet.cell(row, column, cents / 100).number_format = _MONEY
    if history:
        chart = LineChart()
        chart.title = f"Monthly cash flow ({currency})"
        chart.add_data(Reference(sheet, min_col=2, max_col=4, min_row=39, max_row=39 + len(history)), titles_from_data=True)
        chart.set_categories(Reference(sheet, min_col=1, min_row=40, max_row=39 + len(history)))
        chart.width = 17
        sheet.add_chart(chart, "D22")


def make_report(entries: list[Entry], ledger_id: str, currency: str) -> Workbook:
    workbook = Workbook()
    dashboard = workbook.active
    dashboard.title = "Dashboard"
    _transaction_sheet(workbook, entries, "expense")
    _transaction_sheet(workbook, entries, "income")
    summary = workbook.create_sheet("Monthly Summary")
    headers = ("Month", "Currency", "Opening balance", "Income", "Spending", "Net change",
               "Closing balance", "Savings rate", "Transactions")
    _append(summary, headers)
    currency = currency.strip().upper()
    current_month = date.today().isoformat()[:7]
    groups = month_groups(entries)
    for code in {currency, *(item.currency for item in entries)}:
        groups.setdefault((current_month, code), [])
    balances = monthly_balances(groups)
    for (month, code), value in reversed(balances.items()):
        _append(summary, (month, code, value.opening / 100, value.income / 100, value.spending / 100,
                          value.net_change / 100, value.closing / 100,
                          value.net_change / value.income if value.income else None, len(groups[(month, code)])))
        for column in range(3, 8):
            summary.cell(summary.max_row, column).number_format = _MONEY
        summary.cell(summary.max_row, 8).number_format = "0.0%"
    _table(summary, "MonthlySummary", headers, (14, 13, 19, 18, 18, 18, 19, 16, 16))
    _dashboard(dashboard, groups, balances, currency, current_month)
    marker = workbook.create_sheet(_MARKER)
    marker["A1"] = ledger_id
    marker.sheet_state = "veryHidden"
    return workbook


def sync_report(ledger: Ledger, destination: Path, currency: str, *, force: bool = False) -> bool:
    """Serialize exports across app instances, replacing the target only after a full write."""
    destination = destination.resolve()
    with ledger.connect() as db:
        db.execute("BEGIN IMMEDIATE")
        metadata = dict(db.execute("SELECT key, value FROM metadata"))
        if destination.exists():
            if not force and metadata["revision"] == metadata["exported"] and (
                metadata.get("destination") == str(destination)
                and metadata.get("report_month") == date.today().isoformat()[:7]
                and metadata.get("report_currency") == currency
                and metadata.get("report_version") == _REPORT_VERSION
            ):
                return False
            try:
                previous = load_workbook(destination, read_only=True)
                try:
                    owned = _MARKER in previous and previous[_MARKER]["A1"].value == metadata["ledger_id"]
                finally:
                    previous.close()
            except PermissionError:
                raise
            except Exception as exc:
                raise LedgerError(f"Cannot verify ownership of {destination.name}. Move it aside and retry.") from exc
            if not owned:
                raise LedgerError(f"{destination.name} belongs to another workbook. Move it aside or choose another path.")
        destination.parent.mkdir(parents=True, exist_ok=True)
        report = make_report(ledger.all_from(db), metadata["ledger_id"], currency)
        temp = None
        try:
            with NamedTemporaryFile(dir=destination.parent, suffix=".xlsx", delete=False) as handle:
                temp = Path(handle.name)
            report.save(temp)
            with temp.open("r+b") as handle:
                os.fsync(handle.fileno())
            os.replace(temp, destination)
        finally:
            report.close()
            if temp is not None:
                temp.unlink(missing_ok=True)
        db.executemany("INSERT OR REPLACE INTO metadata VALUES (?, ?)", (
            ("exported", metadata["revision"]), ("destination", str(destination)),
            ("report_month", date.today().isoformat()[:7]), ("report_currency", currency),
            ("report_version", _REPORT_VERSION),
        ))
        return True

