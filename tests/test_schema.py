import pytest
from pydantic import ValidationError

from expense_tracker.schema import Category, Expense, LLMExpense, PaymentMethod, ollama_json_schema


def make(**kw):
    base = dict(date="2026-11-05", category="TRANSPORT", amount=230, raw_note="x")
    base.update(kw)
    return Expense(**base)


def test_defaults():
    e = make()
    assert e.currency == "INR"
    assert e.merchant == "N/A"
    assert e.payment_method == "Unknown"


def test_amount_cleanup_and_abs():
    assert make(amount="1,250.50").amount == 1250.5
    assert make(amount="₹450").amount == 450.0
    assert make(amount=-99).amount == 99.0


def test_zero_amount_rejected():
    with pytest.raises(ValidationError):
        make(amount=0)


def test_bad_date_rejected():
    with pytest.raises(ValidationError):
        make(date="2026-13-45")
    with pytest.raises(ValidationError):
        make(date="yesterday")


def test_category_aliases_and_fallback():
    assert make(category="food").category == "FOOD"
    assert make(category="Food and Dining").category == "FOOD"
    assert make(category="totally unknown").category == "OTHERS"


def test_payment_aliases():
    assert make(payment_method="GPay").payment_method == "UPI"
    assert make(payment_method="credit").payment_method == "Credit Card"
    assert make(payment_method=None).payment_method == "Unknown"
    assert make(payment_method="bitcoin").payment_method == "Unknown"


def test_currency_symbols():
    assert make(currency="₹").currency == "INR"
    assert make(currency="usd").currency == "USD"
    assert make(currency="").currency == "INR"


def test_blank_merchant():
    assert make(merchant="  ").merchant == "N/A"
    assert make(merchant=None).merchant == "N/A"


def test_ollama_schema_matches_enums():
    s = ollama_json_schema()
    cats = s["$defs"]["Category"]["enum"]
    assert cats == [c.value for c in Category]
    pays = s["$defs"]["PaymentMethod"]["enum"]
    assert pays == [p.value for p in PaymentMethod]
    assert set(s["required"]) == {"expenses"}
    assert s["properties"]["expenses"]["type"] == "array"
    assert set(s["$defs"]["LLMExpense"]["required"]) == {"category", "amount"}


def test_llm_model_rejects_missing_or_bad_amount():
    with pytest.raises(ValidationError):
        LLMExpense(category="OTHERS")
    with pytest.raises(ValueError):
        LLMExpense(category="OTHERS", amount="abc")


def test_llm_model_normalizes_near_misses():
    # Regression: lowercase "cash" and "1,250" once failed LLM-side validation.
    m = LLMExpense(category="food", amount="1,250", payment_method="cash", currency="₹", merchant=" ")
    assert m.category == Category.FOOD
    assert m.amount == 1250.0
    assert m.payment_method == PaymentMethod.CASH
    assert m.currency == "INR"
    assert m.merchant == "N/A"
