"""Turn local model output into validated entries and read-only queries."""
from __future__ import annotations

import ast
import operator
import re
from collections import Counter
from datetime import date
from decimal import Decimal, DecimalException, localcontext
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, ValidationError

from .config import Settings
from .dates import _NOTE_PATTERNS, parse_date_phrase, resolve_item_date
from .ledger import Entry, LedgerError, as_cents, money
from .ledger_report import breakdown, totals
from .llm import OllamaClient
from .schema import Category, PaymentMethod


class Draft(BaseModel):
    model_config = ConfigDict(extra="forbid")

    kind: Literal["expense", "income"]
    amount: Decimal
    date_text: str
    currency: str
    name: str
    category: Category
    payment: PaymentMethod


class NoteResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    entries: list[Draft] = Field(max_length=20)
    clarification: str


_NOTE_PROMPT = """Extract personal transactions from the user's note as JSON.
Use entries for the transactions, clarification for a short question if any transaction is ambiguous.
Return one entry per stated amount, in the input order, and never sum them. Income is money
received, and expense is money paid. A note can contain both. Skip hypothetical transactions.
For each entry, copy the date phrase verbatim into date_text, or use an empty string if unstated.
Copy the merchant or income source into name, or use N/A. Use the requested default currency
when the note gives no currency. Use the ISO code for explicitly stated currencies.
Use only allowed categories and payment methods. Use OTHERS and Unknown when unspecified.
UPI includes GPay, PhonePe and Paytm. Cab, bus, train and fuel are TRANSPORT. Restaurants and
lunch are FOOD. Use PARKING for parking fees. Use OTHERS and Unknown for income.
If dates are shared, copy the applicable date phrase into each relevant entry.
Never infer missing amounts or invent an entry. If no transactions are clear, return an empty list.
"""


_AMOUNTS = re.compile(r"(?<![\w.])(?:\d+(?:,\d+)*)(?:\.\d+)?(?![\w.])")
_DATE_PATTERN = re.compile(
    "|".join(f"(?:{pattern})" for pattern in _NOTE_PATTERNS)
    + r"|\b(?:today|tonight|this morning|this afternoon|this evening)\b",
    re.I,
)


def _segment_facts(note: str, today: date) -> dict[int, tuple[date | None, str]]:
    segments: dict[int, list[tuple[date | None, str]]] = {}
    for part in re.split(r"\s*\+\s*|[;\n]|\.\s+|\s+and\s+", note, flags=re.I):
        numbers = _AMOUNTS.findall(_DATE_PATTERN.sub(" ", part))
        if len(numbers) != 1:
            continue
        try:
            cents = as_cents(numbers[0])
        except LedgerError:
            continue
        day = parse_date_phrase(part, today)
        methods = {
            PaymentMethod.UPI.value: r"\b(?:upi|gpay|phonepe|paytm)\b",
            PaymentMethod.CASH.value: r"\bcash\b",
            PaymentMethod.CREDIT_CARD.value: r"\bcredit card\b",
            PaymentMethod.DEBIT_CARD.value: r"\bdebit card\b",
            PaymentMethod.NETBANKING.value: r"\bnet\s?banking\b",
        }
        matches = [method for method, pattern in methods.items() if re.search(pattern, part, re.I)]
        segments.setdefault(cents, []).append((day, matches[0] if len(matches) == 1 else ""))
    return {cents: facts[0] for cents, facts in segments.items() if len(facts) == 1}


def parse_note(
    note: str, settings: Settings, client: OllamaClient, today: date | None = None,
) -> tuple[list[Entry], str]:
    note = note.strip()
    if not 1 <= len(note) <= 3000:
        raise LedgerError("Enter a transaction note between 1 and 3,000 characters.")
    today = today or date.today()
    try:
        result = NoteResult.model_validate(client.chat_json(
            _NOTE_PROMPT + f"\nDefault currency: {settings.default_currency}.", note, NoteResult.model_json_schema(),
        ))
    except ValidationError as exc:
        raise LedgerError("Ollama returned an incomplete result. Try a shorter note or use Manual entry.") from exc

    entries: list[Entry] = []
    review = result.clarification.strip()[:350]
    segments = _segment_facts(note, today)
    remaining_note = note.casefold()
    explicit_dates = {match.group(0) for match in _DATE_PATTERN.finditer(remaining_note)}
    stated_dates = set()
    for draft in result.entries:
        phrase = draft.date_text.strip().casefold()
        cents = as_cents(draft.amount)
        segment_date, segment_payment = segments.get(cents, (None, ""))
        if phrase and phrase in note.casefold() and parse_date_phrase(phrase, today):
            stated_dates.add(phrase)
        elif phrase:
            review = review or "Check the dates. Ollama returned a date phrase that could not be verified."
        if draft.name.strip().casefold() not in {"n/a", ""} and draft.name.strip().casefold() not in note.casefold():
            review = review or "Check the merchant or income source."
        if draft.kind == "income" and draft.name.strip().casefold() in {"n/a", ""}:
            review = review or "Enter an income source before saving."
        code = draft.currency.strip().upper() or settings.default_currency.strip().upper()
        aliases = {"INR": r"₹|\brupees?\b|\brs\b", "USD": r"\$|\bus dollars?\b",
                   "EUR": r"€|\beuros?\b", "GBP": r"£|\bpounds?\b"}
        currency_mentioned = re.search(rf"\b{re.escape(code)}\b|{aliases.get(code, '(?!)')}", note, re.I)
        if code != settings.default_currency.strip().upper() and not currency_mentioned:
            review = review or "Check the currency. The note did not explicitly mention it."
        entries.append(Entry(
            draft.kind, (segment_date or resolve_item_date(phrase, note, today)).isoformat(), cents,
            code, draft.name.strip() or "N/A",
            draft.category.value if draft.kind == "expense" else "",
            (segment_payment or draft.payment.value) if draft.kind == "expense" else "", note,
        ))

    for phrase in sorted(stated_dates | explicit_dates, key=len, reverse=True):
        remaining_note = remaining_note.replace(phrase, " ")
    numbers = _AMOUNTS.findall(remaining_note)
    amounts: Counter[int] = Counter()
    for value in numbers:
        candidate = Decimal(value.replace(",", ""))
        if candidate and candidate <= 1_000_000_000_000:
            try:
                amounts[as_cents(candidate)] += 1
            except LedgerError:
                review = review or "Check the amounts. Your note contains a number with more than two decimal places."
    if amounts != Counter(item.cents for item in entries):
        review = review or "Check the amounts and number of entries against your original note."
    for code, token in (("USD", "$"), ("EUR", "€"), ("GBP", "£"), ("INR", "₹")):
        if token in note and not any(item.currency == code for item in entries):
            review = review or "Check the currency against the symbols in your note."
    if len(explicit_dates | stated_dates) > 1 and any(
        segments.get(item.cents, (None, ""))[0] is None for item in entries
    ):
        review = review or "Check each transaction date. Your note contains multiple dates."
    if re.search(r"\b(?:if|would|could|might|maybe|estimate|expected|will pay|will receive)\b|\?", note, re.I):
        review = review or "Confirm these are completed transactions before saving."
    if not entries:
        raise LedgerError(review or "No completed transactions were found. Include an amount and what it was for.")
    return entries, review


class Query(BaseModel):
    model_config = ConfigDict(extra="forbid")

    action: Literal["totals", "compare", "top", "breakdown", "savings", "calculator", "unsupported"]
    kind: Literal["expense", "income", "all"] = "all"
    start: date | None = None
    end: date | None = None
    compare_start: date | None = None
    compare_end: date | None = None
    currency: str = ""
    category: Category | None = None
    payment: PaymentMethod | None = None
    name: str = Field(default="", max_length=200)
    group: Literal["category", "name", "payment"] = "category"
    limit: int = Field(default=5, ge=1, le=30)
    percent: Decimal | None = None
    expression: str = Field(default="", max_length=150)


_OPERATORS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv}


def calculate(expression: str) -> Decimal:
    if not 1 <= len(expression) <= 150 or not re.fullmatch(r"[\d.()+*/\s-]+", expression):
        raise LedgerError("Use numbers, parentheses, and +, −, ×, or ÷.")
    try:
        tree = ast.parse(expression.strip(), mode="eval")
        if sum(1 for _ in ast.walk(tree)) > 45:
            raise LedgerError("That calculation is too long. Split it into smaller parts.")

        def value(node: ast.AST) -> Decimal:
            if isinstance(node, ast.Constant) and type(node.value) in (float, int):
                source = ast.get_source_segment(expression.strip(), node)
                return Decimal(source)
            if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
                sign = -1 if isinstance(node.op, ast.USub) else 1
                return sign * value(node.operand)
            if isinstance(node, ast.BinOp) and type(node.op) in _OPERATORS:
                return _OPERATORS[type(node.op)](value(node.left), value(node.right))
            raise LedgerError("Use numbers, parentheses, and +, −, ×, or ÷.")

        with localcontext() as context:
            context.prec = 28
            answer = value(tree.body)
    except (SyntaxError, DecimalException) as exc:
        raise LedgerError("That calculation is invalid. Check the expression and any division by zero.") from exc
    return answer


def _calculation_answer(expression: str) -> str:
    result = calculate(expression)
    return f"{expression} = {result:,.12f}".rstrip("0").rstrip(".")


def answer_question(question: str, entries: list[Entry], currency: str, client: OllamaClient) -> str:
    question = question.strip()
    if not 1 <= len(question) <= 1500:
        raise LedgerError("Enter a question between 1 and 1,500 characters.")
    plain = question.replace("×", "*").replace("÷", "/").replace("−", "-").replace(",", "").removesuffix("?").strip()
    if re.fullmatch(r"[\d.()+*/\s-]+", plain):
        return _calculation_answer(plain)
    percent = re.fullmatch(r"(?:what(?:'s| is)\s+)?([\d.]+)%\s+of\s+(?:₹|\$|INR\s*|USD\s*)?([\d.]+)", plain, re.I)
    if percent:
        return _calculation_answer(f"{percent[2]} * {percent[1]} / 100")
    today = date.today()
    prompt = (
        f"Convert the question into one supported ledger query. Today is {today:%Y-%m-%d}. "
        f"The default currency is {currency}. All dates are inclusive. "
        "Resolve time periods into ISO dates, including comparison periods. Leave dates null only for all history. "
        "Use totals for income, spending or net, compare for two periods, top for biggest transactions, "
        "breakdown for groups, savings for a hypothetical reduction in filtered spending, "
        "and calculator for ordinary arithmetic. The calculator accepts only decimal numbers, "
        "parentheses, +, -, *, /. Use unsupported for requests outside those operations. "
        "Set kind to expense for spending or income for earnings. "
        "For a percentage reduction, percent is a number from 0 to 100. "
        "Omit category, payment, and name when the question gives no such filter. "
        "For single-day queries set start and end to the same date. For example, How much did I spend yesterday? is a totals query with kind expense and no category, payment, or name filter."
    )
    schema = Query.model_json_schema()
    reduction = re.search(r"(\d+(?:\.\d+)?)\s*%", question)
    if reduction and re.search(r"\b(?:save|saving|less|reduce|reduction|cut)\b", question, re.I) and re.search(
        r"\b(?:spend|spending|expenses?|food|dining|transport|groceries|grocery|entertainment|parking)\b", question, re.I
    ):
        schema["properties"]["action"] = {"const": "savings"}
        prompt += " This question is a savings query about recorded expenses. Use action savings and extract the date and expense filters."
    try:
        plan = Query.model_validate(client.chat_json(prompt, question, schema))
    except ValidationError as exc:
        raise LedgerError("Ollama could not interpret the question. Try one metric and an explicit date range.") from exc
    if schema["properties"]["action"].get("const") == "savings":
        plan = plan.model_copy(update={"action": "savings", "percent": Decimal(reduction[1])})
    return run_query(plan, entries, currency)


def run_query(plan: Query, entries: list[Entry], currency: str) -> str:
    if plan.action == "calculator":
        return _calculation_answer(plan.expression)
    if plan.action == "unsupported":
        raise LedgerError("Ask about transactions, totals, breakdowns, comparisons, or arithmetic.")
    code = plan.currency.strip().upper() or currency
    if not re.fullmatch(r"[A-Z]{3}", code):
        raise LedgerError("Specify a three-letter currency, such as INR.")
    kind = "expense" if plan.action in {"savings", "breakdown"} else plan.kind

    def select(start: date | None, end: date | None) -> tuple[list[Entry], str]:
        if start and end and start > end:
            raise LedgerError("The start date must come before the end date.")
        first = start.isoformat() if start else "0001-01-01"
        last = end.isoformat() if end else "9999-12-31"
        rows = [
            item for item in entries if item.currency == code and first <= item.day <= last
            and (kind == "all" or item.kind == kind)
            and (plan.category is None or item.category == plan.category.value)
            and (plan.payment is None or item.payment == plan.payment.value)
            and plan.name.casefold() in item.name.casefold()
        ]
        period = f"{start or 'earliest'} through {end or 'latest'}" if start or end else "All recorded dates"
        return rows, period

    rows, period = select(plan.start, plan.end)
    filters = [value for value in (kind if kind != "all" else "", plan.category, plan.payment, plan.name) if value]
    scope = f"{period} · {code} · {len(rows)} matching transactions"
    if filters:
        scope += " · " + ", ".join(value.value if hasattr(value, "value") else value for value in filters)

    income, spent = totals(rows)
    if plan.action == "top":
        matches = sorted(rows, key=lambda item: -item.cents)[:plan.limit]
        result = "\n".join(f"{item.day}    {item.name}    {money(item.cents)} {code}" for item in matches)
    elif plan.action == "breakdown":
        result = "\n".join(f"{label}    {money(cents)} {code}" for label, cents in breakdown(rows, plan.group)[:plan.limit])
    elif plan.action == "savings":
        if plan.percent is None or not plan.percent.is_finite() or not 0 <= plan.percent <= 100:
            raise LedgerError("Specify a spending reduction between 0% and 100%.")
        saving = Decimal(spent) * plan.percent / 10000
        result = f"A {plan.percent}% reduction on {money(spent)} {code} would save {saving:,.2f} {code}."
    elif plan.action == "compare":
        if not all((plan.start, plan.end, plan.compare_start, plan.compare_end)):
            raise LedgerError("Specify both date ranges for the comparison.")
        other, other_period = select(plan.compare_start, plan.compare_end)
        old_income, old_spent = totals(other)
        if kind == "income":
            current, previous, label = income, old_income, "Income"
        elif kind == "expense":
            current, previous, label = spent, old_spent, "Spending"
        else:
            current, previous, label = income - spent, old_income - old_spent, "Net balance"
        result = (
            f"{label}: {money(current)} {code}\n"
            f"{other_period}: {money(previous)} {code} ({len(other)} transactions)\n"
            f"Difference (first period minus second): {money(current - previous)} {code}"
        )
    elif kind == "income":
        result = f"Income: {money(income)} {code}"
    elif kind == "expense":
        result = f"Spending: {money(spent)} {code}"
    else:
        result = f"Income: {money(income)} {code}\nSpending: {money(spent)} {code}\nNet balance: {money(income - spent)} {code}"
    return scope + "\n\n" + (result or "No matching transactions.")

