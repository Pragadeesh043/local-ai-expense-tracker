from dataclasses import replace
from datetime import date
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook

from expense_tracker.config import Settings
from expense_tracker.ledger import Entry, Ledger, LedgerError
from expense_tracker.ledger_assistant import Query, answer_question, calculate, parse_note, run_query
from expense_tracker.ledger_report import make_report, month_groups, monthly_balances, sync_report


def expense(name="Cab", *, cents=49000, currency="INR"):
    return Entry("expense", "2026-11-21", cents, currency, name, "TRANSPORT", "UPI")


def test_locked_export_retains_ledger_and_retries(tmp_path, monkeypatch):
    ledger = Ledger(tmp_path / "test.sqlite3")
    report = tmp_path / "report.xlsx"
    ledger.add([expense()])
    assert sync_report(ledger, report, "INR")
    original = report.read_bytes()
    ids = ledger.add([expense("=1+2")])

    def locked(*_args):
        raise PermissionError("Excel has the file open")

    with monkeypatch.context() as patch:
        patch.setattr("expense_tracker.ledger_report.os.replace", locked)
        with pytest.raises(PermissionError):
            sync_report(ledger, report, "INR")
    assert report.read_bytes() == original
    assert len(ledger.entries()) == 2
    assert list(tmp_path.glob("*.xlsx")) == [report]
    assert sync_report(ledger, report, "INR")
    assert not sync_report(ledger, report, "INR")
    workbook = load_workbook(report)
    try:
        assert workbook["Expenses"]["E2"].value == "=1+2"
        assert workbook["Expenses"]["E2"].data_type == "s"
        assert "ExpenseLedger" in workbook["Expenses"].tables
    finally:
        workbook.close()
    ledger.remove(ids)
    assert sync_report(ledger, report, "INR")
    assert len(ledger.entries()) == 1


def test_never_overwrite_unowned_workbook(tmp_path):
    ledger = Ledger(tmp_path / "test.sqlite3")
    report = tmp_path / "report.xlsx"
    Workbook().save(report)
    original = report.read_bytes()
    with pytest.raises(LedgerError, match="belongs to another workbook"):
        sync_report(ledger, report, "INR")
    assert report.read_bytes() == original


def test_import_rolls_back_on_bad_rows_and_imports_only_once(tmp_path):
    ledger = Ledger(tmp_path / "test.sqlite3")
    source = tmp_path / "original.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Expenses"
    sheet.append(["Date", "Category", "Amount", "Currency", "Merchant", "Payment Method", "Raw Note"])
    sheet.append(["21/11/2026", "Food", 12.25, "INR", "Cafe", "Cash", ""])
    sheet.append(["bad date", "Food", 1, "INR", "Cafe", "Cash", ""])
    workbook.save(source)
    with pytest.raises(LedgerError, match="Expenses row 3"):
        ledger.import_once(source, "Expenses", "INR")
    assert ledger.entries() == []
    sheet.delete_rows(3)
    sheet.append(["TOTAL (2026-11)", None, "=SUM(C2:C2)"])
    workbook.save(source)
    assert ledger.import_once(source, "Expenses", "INR") == 1
    assert ledger.import_once(source, "Expenses", "INR") == 0
    assert ledger.entries()[0].cents == 1225
    assert ledger.entries()[0].day == "2026-11-21"


def test_stale_edit_and_partial_undo_roll_back(tmp_path):
    ledger = Ledger(tmp_path / "test.sqlite3")
    ids = ledger.add([expense(), expense("Bus")])
    before = ledger.entries()[0]
    ledger.update(before, replace(before, name="Metro"))
    with pytest.raises(LedgerError, match="changed"):
        ledger.update(before, replace(before, name="Other"))
    with pytest.raises(LedgerError, match="changed"):
        ledger.remove([ids[0], 999])
    assert {item.name for item in ledger.entries()} == {"Cab", "Metro"}


class FakeModel:
    def __init__(self, reply):
        self.reply = reply

    def chat_json(self, *_args):
        return self.reply


def test_mixed_note_checks_amounts_currency_and_dates():
    model = FakeModel({
        "entries": [
            {"kind": "income", "name": "salary", "amount": "75000", "currency": "INR",
             "date_text": "yesterday", "category": "OTHERS", "payment": "Unknown"},
            {"kind": "expense", "name": "cab", "amount": "490", "currency": "INR",
             "date_text": "yesterday", "category": "TRANSPORT", "payment": "Unknown"},
        ],
        "clarification": "",
    })
    note = "salary 75000 today + cab 490 UPI yesterday"
    today = date(2026, 11, 22)
    entries, warning = parse_note(note, Settings(), model, today)
    assert not warning
    assert [(item.kind, item.day, item.cents) for item in entries] == [
        ("income", "2026-11-22", 7500000), ("expense", "2026-11-21", 49000)
    ]
    assert entries[1].payment == "UPI"
    assert "check the amounts" in parse_note(note + " + coffee 50", Settings(), model, today)[1].lower()
    assert "currency" in parse_note(note.replace("490", "$490"), Settings(), model, today)[1].lower()


def test_assistant_sums_one_currency_and_restricts_calculator():
    entries = [expense(cents=49000), expense(cents=100, currency="USD")]
    query = Query(action="savings", percent=Decimal(20))
    assert "98.00 INR" in run_query(query, entries, "INR")
    assert calculate("0.1 + 0.2") == Decimal("0.3")
    for expression in ("1 / 0", "2**2000", "open('file')"):
        with pytest.raises(LedgerError):
            calculate(expression)



def test_opening_balance_carries_forward_across_empty_months_per_currency():
    entries = [
        Entry("income", "2026-09-12", 50000, "INR", "Salary"),
        expense(cents=12500),
        Entry("income", "2026-10-01", 100000, "USD", "Freelance"),
        Entry("income", "2026-12-05", 20000, "INR", "Salary"),
    ]
    groups = month_groups(entries)
    groups[("2026-10", "INR")] = []
    balances = monthly_balances(groups)
    assert balances[("2026-10", "INR")].opening == balances[("2026-10", "INR")].closing == 50000
    assert balances[("2026-10", "USD")].opening == 0
    assert balances[("2026-11", "INR")].opening == 50000
    assert balances[("2026-11", "INR")].closing == 37500
    assert balances[("2026-12", "INR")].closing == 57500

    report = make_report(entries, "test", "INR")
    try:
        rows = list(report["Monthly Summary"].values)
        assert rows[0][2:7] == ("Opening balance", "Income", "Spending", "Net change", "Closing balance")
        november = next(row for row in rows if row[:2] == ("2026-11", "INR"))
        assert november[2:7] == (500, 0, 125, -125, 375)
    finally:
        report.close()

    question = Query(action="totals", start=date(2026, 11, 1), end=date(2026, 11, 30))
    reply = run_query(question, entries, "INR").splitlines()
    assert "Opening balance: 500.00 INR" in reply
    assert "Net change: -125.00 INR" in reply
    assert "Closing balance: 375.00 INR" in reply
    direct = answer_question("What was my closing balance in November 2026?", entries, "INR", client=None)
    assert "Closing balance: 375.00 INR" in direct


def test_previous_report_version_refreshes_without_transaction_change(tmp_path):
    ledger = Ledger(tmp_path / "test.sqlite3")
    report = tmp_path / "report.xlsx"
    assert sync_report(ledger, report, "INR")
    with ledger.connect() as db:
        db.execute("DELETE FROM metadata WHERE key = 'report_version'")
    assert sync_report(ledger, report, "INR")
    assert not sync_report(ledger, report, "INR")
