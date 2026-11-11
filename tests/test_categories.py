import pytest

from expense_tracker.config import Settings
from expense_tracker.extract import extract_expenses
from expense_tracker.schema import Category, Expense, LLMExpense, normalize_category, ollama_json_schema

SEVEN = ["FOOD", "GROCERY", "TRANSPORT", "PARKING", "ENTERTAINMENT", "PERSONAL", "OTHERS"]


def test_enum_is_exactly_the_seven():
    assert [c.value for c in Category] == SEVEN


def test_ollama_schema_enforces_the_seven():
    s = ollama_json_schema()
    assert s["$defs"]["Category"]["enum"] == SEVEN
    assert "Category" in str(s["$defs"]["LLMExpense"]["properties"]["category"])


def test_prompt_lists_all_seven(tmp_path):
    class C:
        def chat_json(self, system, user, schema):
            self.system = system
            return {"expenses": [{"category": "FOOD", "amount": 1}]}

    c = C()
    extract_expenses("x 1", Settings(xlsx_path=tmp_path / "e.xlsx"), client=c)
    for name in SEVEN:
        assert name in c.system
    assert "Food & Dining" not in c.system and "Groceries" not in c.system


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("FOOD", "FOOD"),
        ("food", "FOOD"),
        ("  Parking ", "PARKING"),
        ("Food & Dining", "FOOD"),  # label from earlier versions
        ("Groceries", "GROCERY"),
        ("Utilities", "OTHERS"),
        ("Shopping", "PERSONAL"),
        ("Health", "PERSONAL"),
        ("Subscriptions", "ENTERTAINMENT"),
        ("Transportation", "TRANSPORT"),
        ("Other", "OTHERS"),
        ("something new", "OTHERS"),
        ("", "OTHERS"),
        (None, "OTHERS"),
    ],
)
def test_normalize_category(raw, expected):
    assert normalize_category(raw).value == expected


def test_models_normalize_on_the_way_in():
    assert Expense(date="2026-11-05", category="Food & Dining", amount=1).category == "FOOD"
    assert LLMExpense(category="movies", amount=1).category == Category.ENTERTAINMENT
