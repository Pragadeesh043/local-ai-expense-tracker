from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook
from typer.testing import CliRunner

from expense_tracker import cli
from expense_tracker.config import Settings
from expense_tracker.extract import ExtractionError
from expense_tracker.income import Income, extract_incomes, parse_simple_income
from expense_tracker.sheet import INCOME_HEADERS, append_expense, append_incomes
from expense_tracker.schema import Expense

TODAY = date(2026, 11, 11)


def S(tmp_path, **kw):
    return Settings(xlsx_path=tmp_path / "e.xlsx", **kw)


class NoModel:
    """Fails the test if the model is called."""

    def chat_json(self, *a):
        raise AssertionError("the model should not be called for a simple note")


class FakeClient:
    def __init__(self, *responses):
        self.responses = list(responses)

    def chat_json(self, system, user, schema):
        return self.responses.pop(0)


# ---------- simple parser ----------


@pytest.mark.parametrize(
    "note,source,amount,day,notes",
    [
        ("salary 75000 2026-11-01", "Salary", 75000.0, "2026-11-01", ""),
        ("salary 75000", "Salary", 75000.0, "2026-11-11", ""),
        ("salary ₹75,000", "Salary", 75000.0, "2026-11-11", ""),
        ("freelance project 5000 yesterday", "Freelance project", 5000.0, "2026-11-10", ""),
        ("interest 120.50 on 3 nov", "Interest", 120.5, "2026-11-03", ""),
        ("bonus 1000 diwali gift", "Bonus", 1000.0, "2026-11-11", "diwali gift"),
        ("Salary rs 500 last friday", "Salary", 500.0, "2026-11-06", ""),
    ],
)
def test_simple_parser(note, source, amount, day, notes):
    [i] = parse_simple_income(note, TODAY)
    assert (i.source, i.amount, i.date, i.notes) == (source, amount, day, notes)


@pytest.mark.parametrize(
    "note",
    ["salary 75000 + freelance 5000", "salary 100 and bonus 50", "5000 salary", "salary", "salary 0", "just words"],
)
def test_simple_parser_declines_what_needs_the_model(note):
    assert parse_simple_income(note, TODAY) is None


def test_simple_note_never_calls_the_model(tmp_path):
    [i] = extract_incomes("salary 75000 2026-11-01", S(tmp_path), client=NoModel(), today=TODAY)
    assert i.amount == 75000.0 and i.date == "2026-11-01"


# ---------- LLM path ----------


def test_llm_path_for_multi_item_notes(tmp_path):
    c = FakeClient(
        {"incomes": [
            {"date_text": "2026-11-01", "source": "salary", "amount": 75000},
            {"date_text": "2026-11-01", "source": "freelance", "amount": 5000, "notes": "logo"},
        ]}
    )
    out = extract_incomes("salary 75000 + freelance 5000 on 2026-11-01", S(tmp_path), client=c, today=TODAY)
    assert [(i.source, i.amount, i.date, i.notes) for i in out] == [
        ("Salary", 75000.0, "2026-11-01", ""),
        ("Freelance", 5000.0, "2026-11-01", "logo"),
    ]


def test_llm_path_resolves_dates_in_python(tmp_path):
    c = FakeClient({"incomes": [{"date_text": "today", "source": "Rent", "amount": 9000}]})
    [i] = extract_incomes("got 9000 from rent yesterday", S(tmp_path), client=c, today=TODAY)
    assert i.date == "2026-11-10"  # the model said "today", the note says "yesterday"


def test_llm_repeats_collapsed_and_bad_output_retried(tmp_path):
    one = {"source": "Rent", "amount": 9000}
    c = FakeClient({"incomes": []}, {"incomes": [one, dict(one)]})
    out = extract_incomes("got 9000 from rent", S(tmp_path), client=c, today=TODAY)
    assert len(out) == 1


def test_llm_gives_up(tmp_path):
    c = FakeClient(*[{"incomes": []}] * 3)
    with pytest.raises(ExtractionError):
        extract_incomes("got money from rent", S(tmp_path), client=c, today=TODAY)


def test_empty_note(tmp_path):
    with pytest.raises(ExtractionError):
        extract_incomes("  ", S(tmp_path), client=NoModel(), today=TODAY)


# ---------- Income sheet ----------


def inc(d, source="Salary", amount=100, notes=""):
    return Income(date=d, source=source, amount=amount, notes=notes)


def test_income_sheet_is_created_with_required_headers(tmp_path):
    p = tmp_path / "e.xlsx"
    append_income_rows = append_incomes(p, "Expenses", [inc("2026-11-01", amount=75000)])
    assert append_income_rows == [2]
    ws = load_workbook(p)["Income"]
    assert [c.value for c in ws[1]] == INCOME_HEADERS == ["Date (YYYY-MM-DD)", "Source", "Amount (INR)", "Notes"]
    assert ws["A1"].font.bold and ws.freeze_panes == "A2"
    assert ws["A2"].value.strftime("%Y-%m-%d") == "2026-11-01"
    assert ws["A2"].number_format == "YYYY-MM-DD" and ws["C2"].number_format == "#,##0.00"


def test_income_sheet_exists_after_a_plain_expense_write(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", Expense(date="2026-11-05", category="FOOD", amount=1, raw_note="x"))
    wb = load_workbook(p)
    assert wb.sheetnames == ["Dashboard", "Expenses", "Income"]
    assert [c.value for c in wb["Income"][1]] == INCOME_HEADERS


def test_income_sorted_newest_first_and_existing_sheet_kept(tmp_path):
    p = tmp_path / "e.xlsx"
    append_incomes(p, "Expenses", [inc("2026-11-01", "Salary", 100)])
    append_incomes(p, "Expenses", [inc("2026-11-09", "Bonus", 50), inc("2026-10-15", "Interest", 5)])
    ws = load_workbook(p)["Income"]
    assert [(ws.cell(r, 1).value.strftime("%Y-%m-%d"), ws.cell(r, 2).value) for r in range(2, 5)] == [
        ("2026-11-09", "Bonus"), ("2026-11-01", "Salary"), ("2026-10-15", "Interest"),
    ]
    assert ws.max_row == 4  # no totals or duplicates on the Income sheet


def test_expenses_and_income_do_not_disturb_each_other(tmp_path):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", Expense(date="2026-11-05", category="FOOD", amount=300, raw_note="x"))
    append_incomes(p, "Expenses", [inc("2026-11-01")])
    append_expense(p, "Expenses", Expense(date="2026-11-06", category="PARKING", amount=40, raw_note="x"))
    wb = load_workbook(p)
    assert wb["Income"].max_row == 2
    cats = [wb["Expenses"].cell(r, 2).value for r in range(2, wb["Expenses"].max_row + 1)]
    assert cats == ["PARKING", "FOOD", None]  # two entries and the monthly TOTAL row


# ---------- CLI ----------


def run(*args, **kw):
    return CliRunner().invoke(cli.app, list(args), **kw)


@pytest.fixture
def no_ollama(monkeypatch):
    monkeypatch.setattr(cli, "OllamaClient", lambda *a, **k: NoModel())


def test_cli_add_income_your_example(tmp_path, no_ollama):
    p = tmp_path / "e.xlsx"
    r = run("add-income", "-y", "-f", str(p), "salary 75000 2026-11-01")
    assert r.exit_code == 0, r.output
    assert "Parsed income" in r.output and "75,000.00" in " ".join(r.output.split())
    ws = load_workbook(p)["Income"]
    assert (ws["B2"].value, ws["C2"].value) == ("Salary", 75000)
    assert ws["A2"].value.strftime("%Y-%m-%d") == "2026-11-01"


def test_cli_add_income_asks_and_can_skip(tmp_path, no_ollama):
    p = tmp_path / "e.xlsx"
    r = run("add-income", "-f", str(p), "salary", "75000", input="n\n")
    assert "Skipped" in r.output
    assert not p.exists()


def test_cli_add_income_model_error_is_reported(tmp_path, monkeypatch):
    from expense_tracker.llm import OllamaError

    class Down:
        def chat_json(self, *a):
            raise OllamaError("Cannot reach Ollama")

    monkeypatch.setattr(cli, "OllamaClient", lambda *a, **k: Down())
    r = run("add-income", "-y", "-f", str(tmp_path / "e.xlsx"), "salary 100 + bonus 50")
    assert r.exit_code == 1 and "Cannot reach Ollama" in r.output


def test_cli_list_income(tmp_path, no_ollama):
    p = tmp_path / "e.xlsx"
    run("add-income", "-y", "-f", str(p), "salary 75000 2026-11-01")
    r = run("list", "--income", "-f", str(p))
    assert r.exit_code == 0 and "Salary" in r.output and "75,000.00" in r.output
    assert "No income yet." in run("list", "--income", "-f", str(tmp_path / "none.xlsx")).output
