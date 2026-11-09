"""Strict expense schema. Single source of truth for validation and the Ollama JSON schema."""

from __future__ import annotations

from datetime import date as date_cls
from enum import Enum

from pydantic import BaseModel, ConfigDict, Field, field_validator


class Category(str, Enum):
    FOOD = "Food & Dining"
    TRANSPORT = "Transport"
    UTILITIES = "Utilities"
    GROCERIES = "Groceries"
    SHOPPING = "Shopping"
    ENTERTAINMENT = "Entertainment"
    HEALTH = "Health"
    SUBSCRIPTIONS = "Subscriptions"
    OTHER = "Other"


class PaymentMethod(str, Enum):
    UPI = "UPI"
    CREDIT_CARD = "Credit Card"
    DEBIT_CARD = "Debit Card"
    CASH = "Cash"
    NETBANKING = "NetBanking"
    UNKNOWN = "Unknown"


# Tolerant aliases so a small model's near-miss still lands on a valid enum value.
_CATEGORY_ALIASES = {
    "food": Category.FOOD,
    "dining": Category.FOOD,
    "food and dining": Category.FOOD,
    "restaurant": Category.FOOD,
    "travel": Category.TRANSPORT,
    "grocery": Category.GROCERIES,
    "bills": Category.UTILITIES,
    "utility": Category.UTILITIES,
    "subscription": Category.SUBSCRIPTIONS,
    "medical": Category.HEALTH,
    "healthcare": Category.HEALTH,
    "misc": Category.OTHER,
    "miscellaneous": Category.OTHER,
}

_PAYMENT_ALIASES = {
    "upi": PaymentMethod.UPI,
    "gpay": PaymentMethod.UPI,
    "google pay": PaymentMethod.UPI,
    "phonepe": PaymentMethod.UPI,
    "paytm": PaymentMethod.UPI,
    "credit": PaymentMethod.CREDIT_CARD,
    "credit card": PaymentMethod.CREDIT_CARD,
    "cc": PaymentMethod.CREDIT_CARD,
    "debit": PaymentMethod.DEBIT_CARD,
    "debit card": PaymentMethod.DEBIT_CARD,
    "card": PaymentMethod.UNKNOWN,
    "cash": PaymentMethod.CASH,
    "netbanking": PaymentMethod.NETBANKING,
    "net banking": PaymentMethod.NETBANKING,
    "unknown": PaymentMethod.UNKNOWN,
}


class _Normalized(BaseModel):
    """Shared tolerant normalizers. Applied to both the LLM output model and the final Expense."""

    @field_validator("category", mode="before", check_fields=False)
    @classmethod
    def _norm_category(cls, v: object) -> object:
        if isinstance(v, str):
            key = v.strip().lower().replace("&", "and")
            for cat in Category:
                if key == cat.value.lower().replace("&", "and"):
                    return cat
            return _CATEGORY_ALIASES.get(key, Category.OTHER)
        return v

    @field_validator("payment_method", mode="before", check_fields=False)
    @classmethod
    def _norm_payment(cls, v: object) -> object:
        if v is None:
            return PaymentMethod.UNKNOWN
        if isinstance(v, str):
            return _PAYMENT_ALIASES.get(v.strip().lower(), PaymentMethod.UNKNOWN)
        return v

    @field_validator("amount", mode="before", check_fields=False)
    @classmethod
    def _norm_amount(cls, v: object) -> object:
        if isinstance(v, str):
            cleaned = v.replace(",", "").replace("₹", "").replace("$", "").strip()
            return abs(float(cleaned))
        if isinstance(v, (int, float)):
            return abs(float(v))
        return v

    @field_validator("currency", mode="before", check_fields=False)
    @classmethod
    def _norm_currency(cls, v: object) -> object:
        if v is None or not str(v).strip():
            return "INR"
        s = str(v).strip()
        symbols = {"₹": "INR", "rs": "INR", "rs.": "INR", "$": "USD", "€": "EUR", "£": "GBP"}
        return symbols.get(s.lower(), s.upper())

    @field_validator("merchant", mode="before", check_fields=False)
    @classmethod
    def _norm_merchant(cls, v: object) -> object:
        if v is None or not str(v).strip():
            return "N/A"
        return str(v).strip()


class Expense(_Normalized):
    model_config = ConfigDict(use_enum_values=True, extra="ignore")

    date: str = Field(description="ISO date YYYY-MM-DD")
    category: Category
    amount: float = Field(gt=0, description="Absolute positive amount")
    currency: str = Field(default="INR", min_length=1, max_length=8)
    merchant: str = Field(default="N/A")
    payment_method: PaymentMethod = PaymentMethod.UNKNOWN
    raw_note: str = ""

    @field_validator("date", mode="before")
    @classmethod
    def _check_date(cls, v: object) -> str:
        if isinstance(v, date_cls):
            return v.isoformat()
        s = str(v).strip()
        # Raises ValueError on anything that is not a real calendar date.
        return date_cls.fromisoformat(s).isoformat()


class LLMExpense(_Normalized):
    """What the model is asked to produce. Date and raw_note are handled in code, not by the LLM."""

    model_config = ConfigDict(extra="ignore")

    date_text: str = Field(
        default="",
        description="Date phrase exactly as written by the user (e.g. 'yesterday', 'last Friday', '3 Nov'). Empty if none.",
    )
    category: Category
    amount: float
    currency: str = "INR"
    merchant: str = "N/A"
    payment_method: PaymentMethod = PaymentMethod.UNKNOWN


MAX_ITEMS = 20


class LLMBatch(BaseModel):
    """Top-level model output: one object per expense item in the note."""

    expenses: list[LLMExpense] = Field(min_length=1, max_length=MAX_ITEMS)


def ollama_json_schema() -> dict:
    """JSON schema passed to Ollama's `format` field. Derived from the model so it cannot drift."""
    return LLMBatch.model_json_schema()
