"""Income entries.

Simple notes like "salary 75000 2026-11-01" are parsed directly, with no model call.
Anything else ("freelance 5000 + interest 120", "got 3000 from rent yesterday") goes to Ollama.
"""

from __future__ import annotations

import json
import re
from collections import Counter
from datetime import date as date_cls
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator

from .config import Settings
from .dates import parse_date_phrase, resolve_item_date
from .extract import ExtractionError, coerce_items, count_in_note
from .llm import OllamaClient, OllamaError
from .schema import MAX_ITEMS, NormalizedModel


class Income(NormalizedModel):
    model_config = ConfigDict(extra="ignore")

    date: str = Field(description="ISO date YYYY-MM-DD")
    source: str = Field(min_length=1)
    amount: float = Field(gt=0)
    notes: str = ""

    @field_validator("date", mode="before")
    @classmethod
    def _check_date(cls, v: object) -> str:
        if isinstance(v, date_cls):
            return v.isoformat()
        return date_cls.fromisoformat(str(v).strip()).isoformat()

    @field_validator("source", mode="before")
    @classmethod
    def _clean_source(cls, v: object) -> str:
        s = " ".join(str(v or "").split())
        return s[:1].upper() + s[1:]

    @field_validator("notes", mode="before")
    @classmethod
    def _clean_notes(cls, v: object) -> str:
        return " ".join(str(v or "").split())


class LLMIncome(NormalizedModel):
    model_config = ConfigDict(extra="ignore")

    date_text: str = Field(
        default="", description="Date phrase exactly as written by the user. Empty if none."
    )
    source: str
    amount: float
    notes: str = ""


class LLMIncomeBatch(BaseModel):
    incomes: list[LLMIncome] = Field(min_length=1, max_length=MAX_ITEMS)


# "<source words> <amount> [<date or notes>]". The source has no digits.
_SIMPLE = re.compile(
    r"^(?P<source>[A-Za-z][^\d+&,;]*?)\s+(?:rs\.?\s*|inr\s*|₹\s*)?"
    r"(?P<amount>\d[\d,]*(?:\.\d+)?)(?:\s+(?P<rest>\S.*))?$",
    re.IGNORECASE,
)
_MULTI = re.compile(r"[+&;]|\band\b|,\s", re.IGNORECASE)


def parse_simple_income(note: str, today: date_cls) -> list[Income] | None:
    """Parse "<source> <amount> [date]". Returns None when the note needs the model."""
    note = note.strip()
    if _MULTI.search(note):
        return None
    m = _SIMPLE.match(note)
    if not m:
        return None
    rest = (m.group("rest") or "").strip()
    day = parse_date_phrase(rest, today) if rest else None
    notes = "" if day or not rest else rest
    try:
        return [
            Income(
                date=(day or today).isoformat(),
                source=m.group("source"),
                amount=m.group("amount"),
                notes=notes,
            )
        ]
    except (ValidationError, ValueError):
        return None


_EXAMPLE_NOTE = "salary 75000 + freelance 5000 on 2026-11-01"
_EXAMPLE_JSON = json.dumps(
    {
        "incomes": [
            {"date_text": "2026-11-01", "source": "Salary", "amount": 75000, "notes": ""},
            {"date_text": "2026-11-01", "source": "Freelance", "amount": 5000, "notes": ""},
        ]
    }
)

_SYSTEM = (
    """You extract income entries from a short note. Reply with JSON only, in the form {"incomes": [ ... ]}.

Items:
- Output one object per separate income. "salary 75000 + freelance 5000" is TWO objects.
- Never add amounts together. Never drop an item.
- A date stated once for the whole note applies to every item.

Fields of each object:
- date_text: copy the date phrase from the note exactly (e.g. "yesterday", "2026-11-01", "3 Nov"). Do not work out the date. Use "" if the note has no date.
- source: who or what paid, in a word or two (e.g. "Salary", "Freelance", "Interest", "Rent").
- amount: the positive number received. Plain number, no symbols or commas.
- notes: any extra detail from the note, or "".

Never invent a source or an amount.

Example
Note: """
    + _EXAMPLE_NOTE
    + "\nJSON: "
    + _EXAMPLE_JSON
    + "\n"
)


def _drop_repeats(items: list[Income], note: str) -> list[Income]:
    seen: Counter = Counter()
    out: list[Income] = []
    for i in items:
        key = (i.date, i.source.lower(), i.amount, i.notes.lower())
        seen[key] += 1
        if seen[key] <= max(1, count_in_note(note, i.amount)):
            out.append(i)
    return out


def extract_incomes(
    note: str,
    settings: Settings,
    client: OllamaClient | None = None,
    today: date_cls | None = None,
) -> list[Income]:
    note = note.strip()
    if not note:
        raise ExtractionError("Empty note.")
    today = today or date_cls.today()

    simple = parse_simple_income(note, today)
    if simple is not None:
        return simple

    client = client or OllamaClient(settings.ollama_url, settings.model, settings.timeout_s)
    schema = LLMIncomeBatch.model_json_schema()
    last_err: Exception | None = None
    for _ in range(settings.retries + 1):
        try:
            raw: Any = client.chat_json(_SYSTEM, note, schema)
            out: list[Income] = []
            for item in coerce_items(raw, "incomes"):
                p = LLMIncome.model_validate(item)
                out.append(
                    Income(
                        date=resolve_item_date(p.date_text, note, today).isoformat(),
                        source=p.source,
                        amount=p.amount,
                        notes=p.notes,
                    )
                )
            return _drop_repeats(out, note)
        except OllamaError:
            raise
        except (ValidationError, ValueError) as e:
            last_err = e
    raise ExtractionError(f"Could not parse valid income after retries: {last_err}")
