import json
from datetime import date
from pathlib import Path

import httpx
import pytest
from openpyxl import load_workbook

from expense_tracker.config import Settings
from expense_tracker.extract import ExtractionError, extract_expenses
from expense_tracker.llm import OllamaClient, OllamaError
from expense_tracker.sheet import HEADERS, SheetError, append_expense, read_recent

TODAY = date(2026, 11, 5)


class FakeClient:
    """Stands in for OllamaClient. Returns queued responses, records calls."""

    def __init__(self, *responses):
        self.responses = list(responses)
        self.calls = []

    def chat_json(self, system, user, schema):
        self.calls.append((system, user, schema))
        r = self.responses.pop(0)
        if isinstance(r, Exception):
            raise r
        return r


def settings(tmp_path: Path, **kw) -> Settings:
    return Settings(xlsx_path=tmp_path / "e.xlsx", retries=kw.pop("retries", 2), **kw)


def test_extract_happy_path(tmp_path):
    c = FakeClient(
        {
            "date_text": "yesterday",
            "category": "Food & Dining",
            "amount": 450,
            "currency": "INR",
            "merchant": "Swiggy",
            "payment_method": "UPI",
        }
    )
    [e] = extract_expenses("swiggy 450 upi yesterday", settings(tmp_path), client=c, today=TODAY)
    assert e.date == "2026-11-04"
    assert e.category == "Food & Dining"
    assert e.amount == 450.0
    assert e.merchant == "Swiggy"
    assert e.payment_method == "UPI"
    assert e.raw_note == "swiggy 450 upi yesterday"


def test_extract_no_date_defaults_today(tmp_path):
    c = FakeClient({"category": "Transport", "amount": 120})
    [e] = extract_expenses("auto 120", settings(tmp_path), client=c, today=TODAY)
    assert e.date == "2026-11-05"
    assert e.merchant == "N/A"
    assert e.payment_method == "Unknown"


def test_extract_retries_bad_output_then_succeeds(tmp_path):
    c = FakeClient({"category": "Other"}, {"category": "Other", "amount": 9})  # first lacks amount
    [e] = extract_expenses("thing 9", settings(tmp_path), client=c, today=TODAY)
    assert e.amount == 9.0
    assert len(c.calls) == 2


def test_extract_gives_up_after_retries(tmp_path):
    bad = {"category": "Other", "amount": "not a number"}
    c = FakeClient(bad, bad, bad)
    with pytest.raises(ExtractionError):
        extract_expenses("x 1", settings(tmp_path), client=c, today=TODAY)
    assert len(c.calls) == 3


def test_extract_does_not_retry_connection_errors(tmp_path):
    c = FakeClient(OllamaError("down"))
    with pytest.raises(OllamaError):
        extract_expenses("x 1", settings(tmp_path), client=c, today=TODAY)
    assert len(c.calls) == 1


def test_extract_empty_note(tmp_path):
    with pytest.raises(ExtractionError):
        extract_expenses("   ", settings(tmp_path), client=FakeClient(), today=TODAY)


def test_prompt_contains_enums_and_currency(tmp_path):
    c = FakeClient({"category": "Other", "amount": 1})
    extract_expenses("x 1", settings(tmp_path), client=c, today=TODAY)
    system, user, schema = c.calls[0]
    assert "Food & Dining" in system and "NetBanking" in system and "INR" in system
    assert user == "x 1"
    assert "properties" in schema


# --- HTTP layer, using httpx MockTransport via monkeypatch ---


def _patch_post(monkeypatch, handler):
    transport = httpx.MockTransport(handler)

    def fake_post(url, json=None, timeout=None):
        with httpx.Client(transport=transport) as c:
            return c.post(url, json=json)

    monkeypatch.setattr(httpx, "post", fake_post)


def test_client_sends_schema_and_parses(monkeypatch):
    seen = {}

    def handler(request: httpx.Request):
        seen["body"] = json.loads(request.content)
        seen["path"] = request.url.path
        return httpx.Response(200, json={"message": {"content": json.dumps({"amount": 5})}})

    _patch_post(monkeypatch, handler)
    out = OllamaClient("http://localhost:11434", "llama3.2").chat_json("sys", "usr", {"type": "object"})
    assert out == {"amount": 5}
    assert seen["path"] == "/api/chat"
    assert seen["body"]["format"] == {"type": "object"}
    assert seen["body"]["stream"] is False
    assert seen["body"]["options"]["temperature"] == 0


def test_client_model_missing(monkeypatch):
    _patch_post(monkeypatch, lambda r: httpx.Response(404, json={"error": "model not found"}))
    with pytest.raises(OllamaError, match="ollama pull"):
        OllamaClient("http://localhost:11434", "nope").chat_json("s", "u", {})


def test_client_connection_refused(monkeypatch):
    def handler(request):
        raise httpx.ConnectError("refused")

    _patch_post(monkeypatch, handler)
    with pytest.raises(OllamaError, match="ollama serve"):
        OllamaClient("http://localhost:11434", "m").chat_json("s", "u", {})


def test_client_bad_json_content(monkeypatch):
    _patch_post(monkeypatch, lambda r: httpx.Response(200, json={"message": {"content": "not json"}}))
    with pytest.raises(OllamaError, match="unparseable"):
        OllamaClient("http://localhost:11434", "m").chat_json("s", "u", {})


# --- Sheet ---


def _exp(**kw):
    from expense_tracker.schema import Expense

    base = dict(
        date="2026-11-04", category="Food & Dining", amount=450, merchant="Swiggy",
        payment_method="UPI", raw_note="swiggy 450 upi yesterday",
    )
    base.update(kw)
    return Expense(**base)


def test_sheet_creates_with_headers_and_appends(tmp_path):
    p = tmp_path / "new" / "e.xlsx"
    r1 = append_expense(p, "Expenses", _exp())
    r2 = append_expense(p, "Expenses", _exp(amount=100, merchant="Uber", category="Transport"))
    assert (r1, r2) == (2, 3)

    ws = load_workbook(p)["Expenses"]
    assert [c.value for c in ws[1]] == HEADERS
    assert ws["A2"].value.date() == date(2026, 11, 4)
    assert ws["B2"].value == "Food & Dining"
    assert ws["C2"].value == 450.0
    assert ws["E3"].value == "Uber"
    assert ws["A1"].font.bold is True
    assert ws.freeze_panes == "A2"


def test_sheet_preserves_existing_workbook(tmp_path):
    from openpyxl import Workbook

    p = tmp_path / "e.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    for i, h in enumerate(HEADERS, 1):
        ws.cell(1, i, h)
    ws.cell(2, 1, date(2026, 11, 1))
    ws.cell(2, 2, "old row")
    other = wb.create_sheet("Budget")
    other["A1"] = "keep me"
    wb.save(p)

    row = append_expense(p, "Expenses", _exp())  # dated 2026-11-04, newer, so it goes on top
    assert row == 2
    wb2 = load_workbook(p)
    assert wb2["Expenses"]["B3"].value == "old row"
    assert wb2["Budget"]["A1"].value == "keep me"


def test_sheet_adds_missing_sheet(tmp_path):
    from openpyxl import Workbook

    p = tmp_path / "e.xlsx"
    wb = Workbook()
    wb.active.title = "Other"
    wb.save(p)
    append_expense(p, "Expenses", _exp())
    wb2 = load_workbook(p)
    assert "Other" in wb2.sheetnames and "Expenses" in wb2.sheetnames
    assert wb2["Expenses"]["A1"].value == "Date"


def test_sheet_corrupt_file(tmp_path):
    p = tmp_path / "e.xlsx"
    p.write_text("not a zip")
    with pytest.raises(SheetError, match="not a readable"):
        append_expense(p, "Expenses", _exp())


def test_sheet_locked_file(tmp_path, monkeypatch):
    p = tmp_path / "e.xlsx"
    append_expense(p, "Expenses", _exp())

    from openpyxl.workbook.workbook import Workbook as WB

    def boom(self, *a, **k):
        raise PermissionError("locked")

    monkeypatch.setattr(WB, "save", boom)
    with pytest.raises(SheetError, match="Close it in Excel"):
        append_expense(p, "Expenses", _exp())


def test_read_recent(tmp_path):
    p = tmp_path / "e.xlsx"
    assert read_recent(p, "Expenses") == []
    for i in range(5):
        append_expense(p, "Expenses", _exp(amount=10 + i))
    rows = read_recent(p, "Expenses", n=3)
    assert [r[2] for r in rows] == [12.0, 13.0, 14.0]
