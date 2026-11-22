"""Local transactions, one-time workbook import, and durable change tracking."""
from __future__ import annotations

import re
import sqlite3
from collections.abc import Iterator
from contextlib import closing, contextmanager
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from uuid import uuid4

from openpyxl import load_workbook

from .dates import cell_date, parse_date_phrase
from .schema import PaymentMethod, normalize_category


class LedgerError(ValueError):
    pass


def as_cents(value: object) -> int:
    text = str(value).strip()
    if "," in text and not re.fullmatch(
        r"(?:\d{1,3}(?:,\d{3})+|\d{1,2}(?:,\d{2})+,\d{3})(?:\.\d{1,2})?", text
    ):
        raise LedgerError("Check the comma grouping in the amount.")
    try:
        amount = Decimal(text.replace(",", ""))
        if not amount.is_finite() or not 0 < amount <= 1_000_000_000_000:
            raise LedgerError("Amount must be positive and at most 1 trillion.")
        scaled = amount * 100
        if scaled != scaled.to_integral_value():
            raise LedgerError("Amount must have at most two decimal places.")
        return int(scaled)
    except (InvalidOperation, TypeError) as exc:
        raise LedgerError("Enter a number for the amount.") from exc


def money(cents: int) -> str:
    return f"{Decimal(cents) / 100:,.2f}"


@dataclass(frozen=True)
class Entry:
    kind: str
    day: str
    cents: int
    currency: str
    name: str
    category: str = ""
    payment: str = ""
    note: str = ""
    id: int | None = None

    def __post_init__(self) -> None:
        if self.kind not in {"expense", "income"}:
            raise LedgerError("Choose Expense or Income.")
        try:
            day = date.fromisoformat(self.day)
        except ValueError as exc:
            raise LedgerError("Enter a valid date in YYYY-MM-DD format.") from exc
        if day.isoformat() != self.day or day.year < 1900:
            raise LedgerError("Enter a date from 1900 onward in YYYY-MM-DD format.")
        if type(self.cents) is not int or not 0 < self.cents <= 100_000_000_000_000:
            raise LedgerError("Amount must be positive and at most 1 trillion.")
        if not re.fullmatch(r"[A-Z]{3}", self.currency):
            raise LedgerError("Enter a three-letter currency code such as INR.")
        if not self.name.strip() or len(self.name) > 200 or len(self.note) > 3000:
            raise LedgerError("Enter a merchant or income source (up to 200 characters). Notes can have up to 3,000 characters.")
        if self.kind == "expense" and (
            self.category != normalize_category(self.category).value
            or self.payment not in {method.value for method in PaymentMethod}
        ):
            raise LedgerError("Choose a valid expense category and payment method.")


_COLUMNS = "kind, day, cents, currency, name, category, payment, note"
_TOTAL = re.compile(r"TOTAL \(\d{4}-\d{2}\)")


class Ledger:
    def __init__(self, path: Path):
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as db:
            db.execute("PRAGMA journal_mode = WAL")
            db.executescript("""
                CREATE TABLE IF NOT EXISTS transactions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    kind TEXT NOT NULL CHECK (kind IN ('expense', 'income')),
                    day TEXT NOT NULL,
                    cents INTEGER NOT NULL CHECK (cents > 0),
                    currency TEXT NOT NULL,
                    name TEXT NOT NULL,
                    category TEXT NOT NULL,
                    payment TEXT NOT NULL,
                    note TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS transactions_day ON transactions(day DESC, id DESC);
                CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
                INSERT OR IGNORE INTO metadata VALUES ('revision', '0');
                INSERT OR IGNORE INTO metadata VALUES ('exported', '-1');
            """)
            db.execute("INSERT OR IGNORE INTO metadata VALUES ('ledger_id', ?)", (uuid4().hex,))

    @contextmanager
    def connect(self) -> Iterator[sqlite3.Connection]:
        with closing(sqlite3.connect(self.path, timeout=5)) as connection, connection:
            connection.row_factory = sqlite3.Row
            yield connection

    @staticmethod
    def all_from(db: sqlite3.Connection) -> list[Entry]:
        return [Entry(**dict(row)) for row in db.execute(
            f"SELECT id, {_COLUMNS} FROM transactions ORDER BY day DESC, id DESC"
        )]

    def entries(self) -> list[Entry]:
        with self.connect() as db:
            return self.all_from(db)

    @staticmethod
    def _insert(db: sqlite3.Connection, entries: list[Entry]) -> list[int]:
        return [
            db.execute(
                f"INSERT INTO transactions ({_COLUMNS}) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (item.kind, item.day, item.cents, item.currency, item.name, item.category, item.payment, item.note),
            ).lastrowid
            for item in entries
        ]

    @staticmethod
    def _changed(db: sqlite3.Connection) -> None:
        db.execute("UPDATE metadata SET value = CAST(value AS INTEGER) + 1 WHERE key = 'revision'")

    def add(self, entries: list[Entry]) -> list[int]:
        if not entries:
            raise LedgerError("There are no transactions to save.")
        with self.connect() as db:
            ids = self._insert(db, entries)
            self._changed(db)
            return ids

    def update(self, previous: Entry, item: Entry) -> None:
        with self.connect() as db:
            updated = db.execute("""
                UPDATE transactions SET kind=?, day=?, cents=?, currency=?, name=?, category=?, payment=?, note=?
                WHERE id=? AND kind=? AND day=? AND cents=? AND currency=? AND name=?
                    AND category=? AND payment=? AND note=?
                """, (item.kind, item.day, item.cents, item.currency, item.name, item.category, item.payment, item.note,
                      previous.id, previous.kind, previous.day, previous.cents, previous.currency, previous.name,
                      previous.category, previous.payment, previous.note))
            if updated.rowcount != 1:
                raise LedgerError("This transaction changed in another window. Refresh before editing.")
            self._changed(db)

    def remove(self, ids: list[int]) -> None:
        with self.connect() as db:
            count = sum(db.execute("DELETE FROM transactions WHERE id=?", (item,)).rowcount for item in ids)
            if count != len(ids):
                raise LedgerError("Some transactions changed in another window. Refresh and try again.")
            self._changed(db)

    def import_once(self, source: Path, expense_sheet: str, currency: str) -> int:
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            if db.execute("SELECT 1 FROM metadata WHERE key = 'imported'").fetchone():
                return 0
            if not source.exists():
                return 0
            entries = read_workbook(source, expense_sheet, currency)
            self._insert(db, entries)
            db.execute("INSERT INTO metadata VALUES ('imported', ?)", (str(source),))
            if entries:
                self._changed(db)
            return len(entries)


def read_workbook(source: Path, expense_sheet: str, currency: str) -> list[Entry]:
    """Import all valid rows or abort without changing the ledger."""
    entries: list[Entry] = []
    problems: list[str] = []
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        if expense_sheet not in workbook and "Income" not in workbook:
            raise LedgerError(f"{source.name} has neither {expense_sheet!r} nor an Income sheet.")
        for sheet, kind, width in ((expense_sheet, "expense", 7), ("Income", "income", 4)):
            if sheet not in workbook:
                continue
            heading = next(workbook[sheet].iter_rows(min_row=1, max_row=1, max_col=7, values_only=True), ())
            headers = tuple(str(value or "").strip().casefold() for value in heading)
            if len(headers) < 4 or not headers[0].startswith("date") or not headers[2].startswith("amount") or (
                headers[1] != ("category" if kind == "expense" else "source")
            ):
                problems.append(f"{sheet}: expected Date, {'Category' if kind == 'expense' else 'Source'}, and Amount as the first columns.")
                continue
            has_income_currency = kind == "income" and headers[3] == "currency"
            rows = workbook[sheet].iter_rows(min_row=2, max_col=5 if has_income_currency else width, values_only=True)
            for number, row in enumerate(rows, start=2):
                if all(v is None for v in row) or _TOTAL.fullmatch(str(row[0])):
                    continue
                try:
                    day = cell_date(row[0])
                    if day is None and isinstance(row[0], str) and re.fullmatch(r"\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}", row[0].strip()):
                        day = parse_date_phrase(row[0], date.today())
                    if day is None:
                        raise LedgerError("missing or invalid date")
                    if kind == "expense":
                        payment = next(
                            (p.value for p in PaymentMethod if p.value.casefold() == str(row[5]).strip().casefold()),
                            "Unknown",
                        )
                        item = Entry(kind, day.isoformat(), as_cents(row[2]), str(row[3] or currency).strip().upper(),
                                     str(row[4] or "N/A"), normalize_category(row[1]).value, payment, str(row[6] or ""))
                    else:
                        income_currency = str(row[3] or currency).strip().upper() if has_income_currency else "INR"
                        notes = row[4] if has_income_currency else row[3]
                        item = Entry(kind, day.isoformat(), as_cents(row[2]), income_currency, str(row[1] or "Unknown"),
                                     note=str(notes or ""))
                    entries.append(item)
                except LedgerError as exc:
                    problems.append(f"{sheet} row {number}: {exc}")
    finally:
        workbook.close()
    if problems:
        details = "\n".join(problems[:8])
        suffix = f"\n… plus {len(problems) - 8} more." if len(problems) > 8 else ""
        raise LedgerError(f"Could not import {source.name}. Correct these rows and restart.\n{details}{suffix}")
    return entries

