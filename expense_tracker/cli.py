"""Typer + Rich CLI.

Usage:
    python -m expense_tracker add "swiggy 450 upi yesterday"
    python -m expense_tracker add-income "salary 75000 2026-11-01"
    python -m expense_tracker chat
    python -m expense_tracker list [--income]
    python -m expense_tracker summarize
    python -m expense_tracker check
"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from .config import Settings
from .extract import ExtractionError, extract_expenses
from .income import Income, extract_incomes
from .llm import OllamaClient, OllamaError
from .schema import Expense
from .sheet import (
    HEADERS,
    INCOME,
    INCOME_HEADERS,
    SheetError,
    append_expenses,
    append_incomes,
    read_recent,
    summarize as rebuild,
)

app = typer.Typer(add_completion=False, help="Local expense tracker. Ollama + Excel, fully offline.")
console = Console()


def _settings(model: Optional[str], file: Optional[Path]) -> Settings:
    base = Settings()
    return replace(base, model=model or base.model, xlsx_path=file or base.xlsx_path)


def _preview(items: list[Expense]) -> Table:
    t = Table(show_header=True, box=None, padding=(0, 2))
    for h in ("Date", "Category", "Amount", "Merchant", "Payment"):
        t.add_column(h)
    for e in items:
        t.add_row(e.date, e.category, f"{e.amount:,.2f} {e.currency}", e.merchant, e.payment_method)
    return t


def _preview_income(items: list[Income]) -> Table:
    t = Table(show_header=True, box=None, padding=(0, 2))
    for h in ("Date", "Source", "Amount (INR)", "Notes"):
        t.add_column(h)
    for i in items:
        t.add_row(i.date, i.source, f"{i.amount:,.2f}", i.notes)
    return t


def _where(rows: list[int]) -> str:
    return f"row {rows[0]}" if len(rows) == 1 else "rows " + ", ".join(str(r) for r in rows)


def _process(note: str, s: Settings, client: OllamaClient, confirm: bool) -> bool:
    """Parse one note (one or more expenses) and save it. Returns True if rows were written."""
    try:
        with console.status("Parsing..."):
            items = extract_expenses(note, s, client=client)
    except (OllamaError, ExtractionError) as err:
        console.print(f"[red]Error:[/red] {err}")
        return False

    n = len(items)
    title = "Parsed expense" if n == 1 else f"Parsed {n} expenses"
    console.print(Panel(_preview(items), title=title, expand=False))
    if confirm and not typer.confirm("Save this?" if n == 1 else f"Save these {n} rows?", default=True):
        console.print("[yellow]Skipped.[/yellow]")
        return False

    try:
        rows = append_expenses(
            s.xlsx_path, s.sheet_name, items, s.sort_order, s.default_currency, s.monthly_totals
        )
    except SheetError as err:
        console.print(f"[red]Error:[/red] {err}")
        return False
    console.print(f"[green]Saved[/green] to {s.xlsx_path} ({_where(rows)}).")
    return True


def _process_income(note: str, s: Settings, client: OllamaClient, confirm: bool) -> bool:
    try:
        with console.status("Parsing..."):
            items = extract_incomes(note, s, client=client)
    except (OllamaError, ExtractionError) as err:
        console.print(f"[red]Error:[/red] {err}")
        return False

    n = len(items)
    console.print(Panel(_preview_income(items), title="Parsed income" if n == 1 else f"Parsed {n} incomes", expand=False))
    if confirm and not typer.confirm("Save this?" if n == 1 else f"Save these {n} rows?", default=True):
        console.print("[yellow]Skipped.[/yellow]")
        return False

    try:
        rows = append_incomes(
            s.xlsx_path, s.sheet_name, items, s.sort_order, s.default_currency, s.monthly_totals
        )
    except SheetError as err:
        console.print(f"[red]Error:[/red] {err}")
        return False
    console.print(f"[green]Saved[/green] to the {INCOME} sheet in {s.xlsx_path} ({_where(rows)}).")
    return True


@app.command()
def add(
    note: list[str] = typer.Argument(..., help="Expense in plain words, e.g. swiggy 450 upi yesterday"),
    model: Optional[str] = typer.Option(None, "--model", "-m", help="Ollama model name"),
    file: Optional[Path] = typer.Option(None, "--file", "-f", help="Target .xlsx path"),
    yes: bool = typer.Option(False, "--yes", "-y", help="Save without asking"),
) -> None:
    """Add one or more expenses."""
    s = _settings(model, file)
    client = OllamaClient(s.ollama_url, s.model, s.timeout_s)
    ok = _process(" ".join(note), s, client, confirm=not yes)
    raise typer.Exit(code=0 if ok else 1)


@app.command(name="add-income")
def add_income(
    note: list[str] = typer.Argument(..., help='Income in plain words, e.g. "salary 75000 2026-11-01"'),
    model: Optional[str] = typer.Option(None, "--model", "-m", help="Ollama model name"),
    file: Optional[Path] = typer.Option(None, "--file", "-f", help="Target .xlsx path"),
    yes: bool = typer.Option(False, "--yes", "-y", help="Save without asking"),
) -> None:
    """Add income to the Income sheet.

    Simple notes ("<source> <amount> [date]") are read directly. Others use the Ollama model.
    """
    s = _settings(model, file)
    client = OllamaClient(s.ollama_url, s.model, s.timeout_s)
    ok = _process_income(" ".join(note), s, client, confirm=not yes)
    raise typer.Exit(code=0 if ok else 1)


@app.command()
def chat(
    model: Optional[str] = typer.Option(None, "--model", "-m"),
    file: Optional[Path] = typer.Option(None, "--file", "-f"),
    yes: bool = typer.Option(False, "--yes", "-y", help="Save without asking"),
) -> None:
    """Interactive loop. One expense per line. Start a line with 'income ' to add income. Empty line to quit."""
    s = _settings(model, file)
    client = OllamaClient(s.ollama_url, s.model, s.timeout_s)
    try:
        client.health()
    except OllamaError as err:
        console.print(f"[red]Error:[/red] {err}")
        raise typer.Exit(code=1)

    console.print(f"Model: [bold]{s.model}[/bold]  File: [bold]{s.xlsx_path}[/bold]")
    console.print("Type an expense, or 'income salary 75000'. Empty line to quit.\n")
    while True:
        try:
            note = console.input("[bold cyan]>[/bold cyan] ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if not note or note.lower() in {"q", "quit", "exit"}:
            break
        if note.lower().startswith("income "):
            _process_income(note[7:], s, client, confirm=not yes)
        else:
            _process(note, s, client, confirm=not yes)


@app.command(name="list")
def list_cmd(
    n: int = typer.Option(10, "--n", "-n", help="How many recent rows"),
    file: Optional[Path] = typer.Option(None, "--file", "-f"),
    income: bool = typer.Option(False, "--income", help="Show income instead of expenses"),
) -> None:
    """Show the most recent rows."""
    s = _settings(None, file)
    try:
        rows = read_recent(s.xlsx_path, INCOME if income else s.sheet_name, n)
    except SheetError as err:
        console.print(f"[red]Error:[/red] {err}")
        raise typer.Exit(code=1)
    if not rows:
        console.print("No income yet." if income else "No expenses yet.")
        return

    def day(v) -> str:
        return v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else str(v)

    if income:
        t = Table(*INCOME_HEADERS)
        for r in rows:
            t.add_row(day(r[0]), str(r[1]), f"{r[2]:,.2f}", str(r[3] or ""))
    else:
        t = Table(*HEADERS[:-1])
        for r in rows:
            t.add_row(day(r[0]), str(r[1]), f"{r[2]:,.2f}", str(r[3]), str(r[4]), str(r[5]))
    console.print(t)


@app.command()
def summarize(
    file: Optional[Path] = typer.Option(None, "--file", "-f", help="Target .xlsx path"),
) -> None:
    """Rebuild the monthly TOTAL rows, the Income sheet and the Dashboard from the current data.

    This also runs on every add. Use it after you edit the sheets by hand.
    """
    s = _settings(None, file)
    try:
        res = rebuild(s.xlsx_path, s.sheet_name, s.sort_order, s.default_currency, s.monthly_totals)
    except SheetError as err:
        console.print(f"[red]Error:[/red] {err}")
        raise typer.Exit(code=1)

    parts = [f"Dashboard with {res.charts} monthly chart(s)"]
    if res.totals_on:
        parts.insert(0, "monthly TOTAL rows")
    else:
        parts.insert(0, "no monthly TOTAL rows (EXPENSE_TOTALS is off)")
    console.print(f"[green]Rebuilt[/green] {s.xlsx_path}: " + ", ".join(parts) + ".")
    console.print("Open the file in Excel to see calculated values.")


@app.command()
def check(
    model: Optional[str] = typer.Option(None, "--model", "-m"),
) -> None:
    """Check that Ollama is running and the model is installed."""
    s = _settings(model, None)
    client = OllamaClient(s.ollama_url, s.model, s.timeout_s)
    try:
        names = client.health()
    except OllamaError as err:
        console.print(f"[red]Error:[/red] {err}")
        raise typer.Exit(code=1)
    if not any(n.split(":")[0] == s.model.split(":")[0] for n in names):
        console.print(f"[yellow]Ollama is up, but '{s.model}' is not installed.[/yellow]")
        console.print(f"Run: ollama pull {s.model}")
        raise typer.Exit(code=1)
    console.print(f"[green]OK[/green]. Ollama is up and '{s.model}' is installed.")


if __name__ == "__main__":
    app()
