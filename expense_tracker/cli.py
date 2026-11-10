"""Typer + Rich CLI.

Usage:
    python -m expense_tracker add "swiggy 450 upi yesterday"
    python -m expense_tracker chat
    python -m expense_tracker list
    python -m expense_tracker check
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from .config import Settings
from .extract import ExtractionError, extract_expenses
from .llm import OllamaClient, OllamaError
from .schema import Expense
from .sheet import HEADERS, SheetError, append_expenses, read_recent, remove_totals, summarize as write_summaries

app = typer.Typer(add_completion=False, help="Local expense tracker. Ollama + Excel, fully offline.")
console = Console()


def _settings(model: Optional[str], file: Optional[Path]) -> Settings:
    base = Settings()
    return Settings(
        ollama_url=base.ollama_url,
        model=model or base.model,
        xlsx_path=file or base.xlsx_path,
        sheet_name=base.sheet_name,
        default_currency=base.default_currency,
        sort_order=base.sort_order,
        timeout_s=base.timeout_s,
        retries=base.retries,
    )


def _preview(items: list[Expense]) -> Table:
    t = Table(show_header=True, box=None, padding=(0, 2))
    for h in ("Date", "Category", "Amount", "Merchant", "Payment"):
        t.add_column(h)
    for e in items:
        t.add_row(e.date, e.category, f"{e.amount:,.2f} {e.currency}", e.merchant, e.payment_method)
    return t


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
        rows = append_expenses(s.xlsx_path, s.sheet_name, items, s.sort_order, s.default_currency)
    except SheetError as err:
        console.print(f"[red]Error:[/red] {err}")
        return False
    where = f"row {rows[0]}" if n == 1 else "rows " + ", ".join(str(r) for r in rows)
    console.print(f"[green]Saved[/green] to {s.xlsx_path} ({where}).")
    return True


@app.command()
def add(
    note: list[str] = typer.Argument(..., help="Expense in plain words, e.g. swiggy 450 upi yesterday"),
    model: Optional[str] = typer.Option(None, "--model", "-m", help="Ollama model name"),
    file: Optional[Path] = typer.Option(None, "--file", "-f", help="Target .xlsx path"),
    yes: bool = typer.Option(False, "--yes", "-y", help="Save without asking"),
) -> None:
    """Add one expense."""
    s = _settings(model, file)
    client = OllamaClient(s.ollama_url, s.model, s.timeout_s)
    ok = _process(" ".join(note), s, client, confirm=not yes)
    raise typer.Exit(code=0 if ok else 1)


@app.command()
def chat(
    model: Optional[str] = typer.Option(None, "--model", "-m"),
    file: Optional[Path] = typer.Option(None, "--file", "-f"),
    yes: bool = typer.Option(False, "--yes", "-y", help="Save without asking"),
) -> None:
    """Interactive loop. One expense per line. Empty line or 'q' to quit."""
    s = _settings(model, file)
    client = OllamaClient(s.ollama_url, s.model, s.timeout_s)
    try:
        client.health()
    except OllamaError as err:
        console.print(f"[red]Error:[/red] {err}")
        raise typer.Exit(code=1)

    console.print(f"Model: [bold]{s.model}[/bold]  File: [bold]{s.xlsx_path}[/bold]")
    console.print("Type an expense. Empty line to quit.\n")
    while True:
        try:
            note = console.input("[bold cyan]>[/bold cyan] ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if not note or note.lower() in {"q", "quit", "exit"}:
            break
        _process(note, s, client, confirm=not yes)


@app.command(name="list")
def list_cmd(
    n: int = typer.Option(10, "--n", "-n", help="How many recent rows"),
    file: Optional[Path] = typer.Option(None, "--file", "-f"),
) -> None:
    """Show the most recent rows."""
    s = _settings(None, file)
    try:
        rows = read_recent(s.xlsx_path, s.sheet_name, n)
    except SheetError as err:
        console.print(f"[red]Error:[/red] {err}")
        raise typer.Exit(code=1)
    if not rows:
        console.print("No expenses yet.")
        return
    t = Table(*HEADERS[:-1])
    for r in rows:
        d = r[0].strftime("%Y-%m-%d") if hasattr(r[0], "strftime") else str(r[0])
        t.add_row(d, str(r[1]), f"{r[2]:,.2f}", str(r[3]), str(r[4]), str(r[5]))
    console.print(t)


@app.command()
def summarize(
    file: Optional[Path] = typer.Option(None, "--file", "-f", help="Target .xlsx path"),
    remove: bool = typer.Option(False, "--remove", help="Remove the monthly TOTAL rows instead"),
    no_dashboard: bool = typer.Option(False, "--no-dashboard", help="Only write monthly TOTAL rows"),
) -> None:
    """Add monthly TOTAL rows and a Dashboard sheet with a category chart.

    After the first run, later `add` commands keep the totals and the Dashboard up to date.
    """
    s = _settings(None, file)
    try:
        if remove:
            remove_totals(s.xlsx_path, s.sheet_name, s.sort_order, s.default_currency)
            console.print(f"[green]Removed[/green] monthly TOTAL rows from {s.xlsx_path}.")
            return
        res = write_summaries(
            s.xlsx_path, s.sheet_name, s.sort_order, s.default_currency,
            totals=True, dashboard=not no_dashboard,
        )
    except SheetError as err:
        console.print(f"[red]Error:[/red] {err}")
        raise typer.Exit(code=1)

    if res.months == 0:
        console.print("No dated expenses found, so no monthly totals were written.")
    else:
        console.print(f"[green]Wrote[/green] {res.months} monthly total row group(s) to {s.xlsx_path}.")
    if res.dashboard_written:
        if res.categories:
            console.print(f"[green]Built[/green] the Dashboard sheet with a chart of {res.categories} categories.")
        else:
            console.print(f"Dashboard built, but there are no {s.default_currency} expenses to chart yet.")
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
