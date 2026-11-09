"""Natural language -> list of validated Expenses.

The LLM splits the note into items and extracts fields. Python resolves every date.
"""

from __future__ import annotations

import json
import re
from collections import Counter
from datetime import date
from typing import Any

from pydantic import ValidationError

from .config import Settings
from .dates import resolve_item_date
from .llm import OllamaClient, OllamaError
from .schema import MAX_ITEMS, Category, Expense, LLMExpense, PaymentMethod, ollama_json_schema

_EXAMPLE_NOTE = "cab 490 + lunch 330 paid via cash yesterday"
_EXAMPLE_JSON = json.dumps(
    {
        "expenses": [
            {"date_text": "yesterday", "category": "Transport", "amount": 490,
             "currency": "INR", "merchant": "N/A", "payment_method": "Cash"},
            {"date_text": "yesterday", "category": "Food & Dining", "amount": 330,
             "currency": "INR", "merchant": "N/A", "payment_method": "Cash"},
        ]
    }
)

_SYSTEM = (
    """You extract expenses from a short note. Reply with JSON only, in the form {"expenses": [ ... ]}.

Items:
- Output one object per separate item. "cab 490 + lunch 330" is TWO objects.
- Never add amounts together. Never drop an item.
- "+", "and", "&", commas and "then" often separate items.
- Details stated once for the whole note (date, payment method, currency) apply to every item.

Fields of each object:
- date_text: copy the date phrase from the note exactly (e.g. "yesterday", "last Friday", "3 Nov"). Do not work out the date. Use "" if the note has no date.
- category: one of """
    + str([c.value for c in Category])
    + """
- amount: the positive number paid for that item. Plain number, no symbols or commas.
- currency: ISO code. Use "{currency}" unless the note names another currency.
- merchant: vendor name if given, else "N/A".
- payment_method: one of """
    + str([p.value for p in PaymentMethod])
    + """. Use "Unknown" if not stated.

Rules:
- Never invent a merchant, amount or payment method.
- Swiggy, Zomato, restaurants, cafes, lunch, dinner -> Food & Dining.
- Uber, Ola, cab, metro, petrol, auto, bus -> Transport.
- Electricity, water, gas, internet, mobile recharge -> Utilities.
- Vegetables, kirana, BigBasket, Blinkit, Zepto -> Groceries.
- Netflix, Spotify, YouTube Premium, app plans -> Subscriptions.
- GPay, PhonePe, Paytm -> UPI.

Example
Note: """
    + _EXAMPLE_NOTE
    + "\nJSON: "
    + _EXAMPLE_JSON
    + "\n"
)


class ExtractionError(RuntimeError):
    pass


def _coerce_items(raw: Any) -> list[dict]:
    """Accept {"expenses": [...]}, a bare list, or a single object. Return a list of dicts."""
    if isinstance(raw, list):
        items = raw
    elif isinstance(raw, dict):
        if isinstance(raw.get("expenses"), list):
            items = raw["expenses"]
        elif "amount" in raw:
            items = [raw]
        else:
            raise ValueError("model output has no expenses")
    else:
        raise ValueError("model output is not an object or list")

    if not items:
        raise ValueError("model returned zero expenses")
    if len(items) > MAX_ITEMS:
        raise ValueError(f"model returned {len(items)} expenses; the limit is {MAX_ITEMS}")
    if not all(isinstance(i, dict) for i in items):
        raise ValueError("model returned a non-object item")
    return items


def _count_in_note(note: str, amount: float) -> int:
    nums = re.findall(r"\d+(?:\.\d+)?", note.replace(",", ""))
    return sum(1 for n in nums if float(n) == amount)


def _drop_repeats(items: list[Expense], note: str) -> list[Expense]:
    """Drop items the model repeated.

    Small models sometimes return the same item twice. Identical items are kept only as many
    times as their amount appears in the note, so "coffee 50 + tea 50" keeps both rows but a
    doubled "zomato 700" keeps one.
    """
    seen: Counter = Counter()
    out: list[Expense] = []
    for e in items:
        key = (e.date, e.category, e.amount, e.currency, e.merchant, e.payment_method)
        seen[key] += 1
        if seen[key] <= max(1, _count_in_note(note, e.amount)):
            out.append(e)
    return out


def extract_expenses(
    note: str,
    settings: Settings,
    client: OllamaClient | None = None,
    today: date | None = None,
) -> list[Expense]:
    note = note.strip()
    if not note:
        raise ExtractionError("Empty note.")

    today = today or date.today()
    client = client or OllamaClient(settings.ollama_url, settings.model, settings.timeout_s)
    system = _SYSTEM.replace("{currency}", settings.default_currency)
    schema = ollama_json_schema()

    last_err: Exception | None = None
    for _ in range(settings.retries + 1):
        try:
            raw = client.chat_json(system, note, schema)
            out: list[Expense] = []
            for item in _coerce_items(raw):
                p = LLMExpense.model_validate(item)
                out.append(
                    Expense(
                        date=resolve_item_date(p.date_text, note, today).isoformat(),
                        category=p.category,
                        amount=p.amount,
                        currency=p.currency or settings.default_currency,
                        merchant=p.merchant,
                        payment_method=p.payment_method,
                        raw_note=note,
                    )
                )
            return _drop_repeats(out, note)
        except OllamaError:
            raise  # connection / model problems will not fix themselves on retry
        except (ValidationError, ValueError) as e:
            last_err = e  # model gave bad output; retry

    raise ExtractionError(f"Could not parse valid expenses after retries: {last_err}")
