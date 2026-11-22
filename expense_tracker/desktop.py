"""Windows-friendly desktop interface. Run with python -m expense_tracker.desktop."""
from __future__ import annotations

import argparse
import logging
import os
import queue
import sqlite3
import threading
import tkinter as tk
from collections.abc import Callable
from dataclasses import replace
from datetime import date
from functools import partial
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from urllib.parse import urlsplit
from zipfile import BadZipFile

from .config import Settings
from .ledger import Entry, Ledger, LedgerError, as_cents, money
from .ledger_assistant import answer_question, parse_note
from .ledger_report import breakdown, sync_report, totals
from .llm import OllamaClient, OllamaError
from .schema import Category, PaymentMethod


_BG = "#F2F6F8"
_INK = "#17344C"
_TEAL = "#087E74"
_MUTED = "#526877"
_WHITE = "#FFFFFF"


def _message(error: Exception) -> str:
    if isinstance(error, (LedgerError, OllamaError)):
        return str(error)
    if isinstance(error, PermissionError):
        return "Permission denied. Close the file in Excel if it is open, then try again."
    if isinstance(error, BadZipFile):
        return "This file is not a readable Excel workbook."
    if isinstance(error, sqlite3.OperationalError) and "locked" in str(error).lower():
        return "The local ledger is busy. Wait a moment and try again."
    if isinstance(error, OSError):
        return f"Could not access local storage. {error.strerror or type(error).__name__}"
    return "The operation failed. See app.log in the local data folder for details."


class Desktop:
    def __init__(self, root: tk.Tk, settings: Settings, data: Path, source: Path, report: Path):
        self.root = root
        self.settings = settings
        self.data = data
        self.source = source
        self.report = report
        self.currency = settings.default_currency.strip().upper()
        self.client = OllamaClient(settings.ollama_url, settings.model, settings.timeout_s, think=False)
        self.ledger: Ledger | None = None
        self.entries: list[Entry] = []
        self.undo_ids: list[int] = []
        self.results: queue.Queue = queue.Queue()
        self.busy = False
        self.pending = False
        self.startup_ok = False
        self.status = tk.StringVar(value="Opening the local ledger…")
        self.excel = tk.StringVar(value="Excel report has not been checked.")
        self.month = tk.StringVar(value=date.today().isoformat()[:7])
        self.code = tk.StringVar(value=self.currency)
        self.filter = tk.StringVar()
        self.auto = tk.BooleanVar(value=True)
        root.title("Expense Tracker")
        root.geometry("1030x735")
        root.minsize(760, 640)
        root.configure(bg=_BG)
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.report_callback_exception = self.callback_error
        self._layout()
        root.after(80, self._collect)
        root.after(200, self.start)
        root.after(20_000, self._periodic)

    def callback_error(self, kind, error, trace) -> None:
        logging.error("Desktop interface error", exc_info=(kind, error, trace))
        self.status.set("The interface encountered an error.")
        messagebox.showerror("Expense Tracker", _message(error), parent=self.root)

    def _layout(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10), background=_BG, foreground=_INK)
        style.configure("TFrame", background=_BG)
        style.configure("TLabel", padding=0)
        style.configure("Title.TLabel", font=("Segoe UI", 28, "bold"), foreground=_INK)
        style.configure("Section.TLabel", font=("Segoe UI", 17, "bold"), foreground=_INK)
        style.configure("Muted.TLabel", foreground=_MUTED)
        style.configure("Card.TLabel", background=_WHITE, padding=10, foreground=_INK)
        style.configure("Value.TLabel", background=_WHITE, padding=(10, 0, 10, 12), font=("Segoe UI", 18, "bold"))
        style.configure("Accent.TButton", background=_TEAL, foreground=_WHITE, padding=(14, 8))
        style.map("Accent.TButton", background=[("active", "#086A63"), ("disabled", "#BCD0CE")])
        style.configure("TButton", padding=(9, 6))
        style.configure("TNotebook", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(20, 9))
        style.map("TNotebook.Tab", background=[("selected", _WHITE)])
        style.configure("Treeview", rowheight=28, fieldbackground=_WHITE, background=_WHITE)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

        shell = ttk.Frame(self.root, padding=(24, 14))
        shell.pack(fill="both", expand=True)
        ttk.Label(shell, text="Expense Tracker", style="Title.TLabel").pack(anchor="w")
        ttk.Label(shell, textvariable=self.status, style="Muted.TLabel", wraplength=690).pack(anchor="w", pady=(0, 12))
        self.tabs = ttk.Notebook(shell)
        self.tabs.pack(fill="both", expand=True)
        for name, build in (
            ("Log", self._log_tab), ("Dashboard", self._dash_tab),
            ("Transactions", self._history_tab), ("Ask", self._ask_tab), ("Settings", self._settings_tab),
        ):
            page = ttk.Frame(self.tabs, padding=(12, 12))
            self.tabs.add(page, text=name)
            build(page)
        footer = ttk.Frame(shell)
        footer.pack(fill="x", pady=(9, 0))
        ttk.Label(footer, textvariable=self.excel, style="Muted.TLabel", wraplength=570).pack(side="left")
        ttk.Button(footer, text="Open Excel report", command=self.open_report).pack(side="right")

    @staticmethod
    def _text(parent, height: int, *, readonly: bool = False) -> tk.Text:
        widget = tk.Text(parent, height=height, font=("Segoe UI", 11), relief="solid", bd=1,
                         padx=10, pady=8, wrap="word", bg=_WHITE, fg=_INK, insertbackground=_INK, undo=True)
        if readonly:
            widget.configure(state="disabled")
        return widget

    def _log_tab(self, page) -> None:
        ttk.Label(page, text="Log transactions", style="Section.TLabel").pack(anchor="w")
        ttk.Label(page, text="Example: salary 75000 today + cab 490 UPI yesterday",
                  style="Muted.TLabel", wraplength=640).pack(anchor="w", pady=(3, 8))
        self.note = self._text(page, 5)
        self.note.pack(fill="x")
        self.note.bind("<Control-Return>", lambda _event: (self.log(), "break")[1])
        actions = ttk.Frame(page)
        actions.pack(fill="x", pady=(9, 16))
        self.log_button = ttk.Button(actions, text="Log note", style="Accent.TButton", command=self.log)
        self.log_button.pack(side="left")
        ttk.Button(actions, text="Manual entry", command=self.manual).pack(side="left", padx=6)
        self.undo_button = ttk.Button(actions, text="Undo last save", command=self.undo, state="disabled")
        self.undo_button.pack(side="left")
        ttk.Checkbutton(actions, text="Automatically save verified notes", variable=self.auto).pack(side="right")
        ttk.Label(page, text="Recent transactions", style="Section.TLabel").pack(anchor="w", pady=(0, 5))
        self.recent = self._tree(page, ("Date", "Type", "Merchant / Source", "Amount"), (100, 82, 210, 125), 5)

    def _dash_tab(self, page) -> None:
        toolbar = ttk.Frame(page)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Label(toolbar, text="Monthly overview", style="Section.TLabel").pack(side="left")
        self.codes = ttk.Combobox(toolbar, width=7, textvariable=self.code, state="readonly")
        self.codes.pack(side="right")
        self.months = ttk.Combobox(toolbar, width=11, textvariable=self.month, state="readonly")
        self.months.pack(side="right", padx=8)
        self.codes.bind("<<ComboboxSelected>>", self.refresh_dashboard)
        self.months.bind("<<ComboboxSelected>>", self.refresh_dashboard)

        cards = ttk.Frame(page)
        cards.pack(fill="x", pady=(0, 12))
        self.cards = [tk.StringVar() for _ in range(4)]
        for i, (label, value) in enumerate(zip(("Income", "Spending", "Net balance", "Savings rate"), self.cards)):
            cards.columnconfigure(i, weight=1, uniform="cards")
            ttk.Label(cards, text=label, style="Card.TLabel").grid(row=0, column=i, sticky="ew", padx=3)
            ttk.Label(cards, textvariable=value, style="Value.TLabel").grid(row=1, column=i, sticky="ew", padx=3)
        self.chart_data: list[tuple[str, int]] = []
        self.spending_chart = tk.Canvas(page, height=118, bg=_WHITE, highlightthickness=0)
        self.spending_chart.pack(fill="x", pady=(0, 10))
        self.spending_chart.bind("<Configure>", lambda _event: self.draw_spending())
        self.details = ttk.Notebook(page)
        self.details.pack(fill="both", expand=True)
        self.breakdowns = {}
        for label, field in (("Categories", "category"), ("Merchants", "name"), ("Payment methods", "payment")):
            frame = ttk.Frame(self.details, padding=(6, 8))
            self.details.add(frame, text=label)
            self.breakdowns[field] = self._tree(frame, (label, "Spent", "Share"), (220, 120, 90), 7)
        trend = ttk.Frame(self.details, padding=(6, 8))
        self.details.add(trend, text="Cash flow")
        self.trend = self._tree(trend, ("Month", "Income", "Spending", "Net balance"), (130, 125, 125, 125), 7)

    def _history_tab(self, page) -> None:
        ttk.Label(page, text="Transactions", style="Section.TLabel").pack(anchor="w")
        row = ttk.Frame(page)
        row.pack(fill="x", pady=(8, 10))
        ttk.Label(row, text="Search").pack(side="left")
        ttk.Entry(row, textvariable=self.filter).pack(side="left", padx=8, fill="x", expand=True)
        ttk.Button(row, text="Refresh", command=self.reload).pack(side="right")
        ttk.Button(row, text="Delete", command=self.delete).pack(side="right", padx=4)
        ttk.Button(row, text="Edit", command=self.edit).pack(side="right")
        self.history = self._tree(
            page, ("Date", "Type", "Merchant / Source", "Category", "Payment", "Amount"),
            (95, 75, 155, 123, 96, 112), 11,
        )
        self.history.bind("<Double-1>", lambda event: self.edit() if self.history.identify_row(event.y) else None)
        self.filter.trace_add("write", lambda *_: self.refresh_history())

    def _ask_tab(self, page) -> None:
        ttk.Label(page, text="Ask about your money", style="Section.TLabel").pack(anchor="w")
        ttk.Label(page, text="Try “How much did I spend on food last month?” or “15% of 42500”.",
                  style="Muted.TLabel", wraplength=640).pack(anchor="w", pady=(3, 8))
        self.question = self._text(page, 3)
        self.question.pack(fill="x")
        self.question.bind("<Control-Return>", lambda _event: (self.ask(), "break")[1])
        ttk.Button(page, text="Ask", style="Accent.TButton", command=self.ask).pack(anchor="w", pady=8)
        self.answer = self._text(page, 11, readonly=True)
        self.answer.pack(fill="both", expand=True)

    def _settings_tab(self, page) -> None:
        ttk.Label(page, text="Local setup", style="Section.TLabel").pack(anchor="w")
        for label, value in (
            ("Ollama model", self.settings.model), ("Ollama endpoint", self.settings.ollama_url),
            ("Data folder", str(self.data)), ("Excel report", str(self.report)),
        ):
            row = ttk.Frame(page)
            row.pack(fill="x", pady=6)
            ttk.Label(row, text=label, width=16).pack(side="left", anchor="n")
            ttk.Label(row, text=value, wraplength=570).pack(side="left", fill="x")
        row = ttk.Frame(page)
        row.pack(fill="x", pady=12)
        ttk.Button(row, text="Check Ollama", command=self.check_model).pack(side="left")
        ttk.Button(row, text="Import existing workbook", command=self.choose_import).pack(side="left", padx=8)
        ttk.Button(row, text="Retry Excel update", command=lambda: self.export(force=True)).pack(side="left")

    @staticmethod
    def _tree(parent, columns, widths, height) -> ttk.Treeview:
        body = ttk.Frame(parent)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=columns, show="headings", height=height, selectmode="browse")
        scroll = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        for col, width in zip(columns, widths, strict=True):
            tree.heading(col, text=col, anchor="w")
            tree.column(col, width=width, minwidth=58)
        scroll.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        return tree

    def start(self) -> None:
        def load():
            ledger = Ledger(self.data / "transactions.sqlite3")
            count = ledger.import_once(self.source, self.settings.sheet_name, self.currency)
            return ledger, count, ledger.entries()

        def ready(result):
            self.ledger, count, items = result
            self.startup_ok = True
            self.render(items)
            self.status.set(f"Imported {count} transactions." if count else "Ready")
            self.export()

        self.run("Opening the local ledger…", load, ready)

    def run(self, label: str, work: Callable, done: Callable, failed: Callable | None = None) -> bool:
        if self.busy:
            return False
        self.busy = True
        self.status.set(label)

        def execute():
            try:
                self.results.put((done, failed, work(), None))
            except Exception as exc:
                logging.exception("Desktop operation failed: %s", label)
                self.results.put((done, failed, None, exc))

        threading.Thread(target=execute, daemon=True).start()
        return True

    def _collect(self) -> None:
        try:
            done, failed, value, error = self.results.get_nowait()
        except queue.Empty:
            self.root.after(80, self._collect)
            return
        self.busy = False
        try:
            if error is None:
                done(value)
            elif failed:
                failed(error)
            else:
                self.status.set("Action could not be completed.")
                messagebox.showerror("Expense Tracker", _message(error), parent=self.root)
        finally:
            self.root.after(80, self._collect)

    def _periodic(self) -> None:
        if self.startup_ok and self.pending and not self.busy and not self.root.grab_current():
            self.export(quiet=True)
        self.root.after(20_000, self._periodic)

    def close(self) -> None:
        if self.busy:
            self.root.bell()
            return
        self.root.destroy()

    def export(self, *, force: bool = False, quiet: bool = False) -> None:
        if not self.ledger or self.busy:
            return
        prior = self.status.get()

        def done(_result):
            self.pending = False
            self.excel.set("Excel report is up to date.")
            self.status.set(prior)

        def failed(error):
            self.pending = isinstance(error, (PermissionError, OSError, sqlite3.OperationalError))
            message = "Close Excel to let the app retry." if isinstance(error, PermissionError) else _message(error)
            self.excel.set("Excel update pending. " + message)
            self.status.set(prior)

        self.run(prior if quiet else "Updating Excel report…",
                 partial(sync_report, self.ledger, self.report, self.currency, force=force), done, failed)

    def render(self, entries: list[Entry]) -> None:
        self.entries = entries
        self.refresh_history()
        self.refresh_dashboard()
        self.recent.delete(*self.recent.get_children())
        for item in entries[:7]:
            self.recent.insert("", "end", values=(item.day, item.kind.title(), item.name,
                                                  f"{money(item.cents)} {item.currency}"))

    def refresh_history(self) -> None:
        query = self.filter.get().strip().casefold()
        self.history.delete(*self.history.get_children())
        for item in self.entries:
            searchable = " ".join((item.day, item.kind, item.name, item.category, item.payment, item.note, item.currency))
            if query in searchable.casefold():
                self.history.insert("", "end", iid=str(item.id), values=(
                    item.day, item.kind.title(), item.name, item.category, item.payment,
                    f"{money(item.cents)} {item.currency}",
                ))

    def refresh_dashboard(self, _event=None) -> None:
        months = {date.today().isoformat()[:7], *(item.day[:7] for item in self.entries)}
        self.months["values"] = sorted(months, reverse=True)
        self.codes["values"] = sorted({self.currency, *(item.currency for item in self.entries)})
        currency = self.code.get()
        month = self.month.get()
        rows = [item for item in self.entries if item.currency == currency and item.day.startswith(month)]
        income, spent = totals(rows)
        values = [f"{money(cents)} {currency}" for cents in (income, spent, income - spent)]
        values.append(f"{(income - spent) / income:.1%}" if income else "—")
        for var, value in zip(self.cards, values, strict=True):
            var.set(value)
        for field, tree in self.breakdowns.items():
            tree.delete(*tree.get_children())
            for label, cents in breakdown(rows, field):
                tree.insert("", "end", values=(label, money(cents), f"{cents / spent:.1%}"))
        self.chart_data = breakdown(rows, "category")[:3]
        self.draw_spending()
        self.trend.delete(*self.trend.get_children())
        groups: dict[str, list[Entry]] = {}
        for entry in self.entries:
            if entry.currency == currency:
                groups.setdefault(entry.day[:7], []).append(entry)
        for key, items in sorted(groups.items(), reverse=True):
            received, paid = totals(items)
            self.trend.insert("", "end", values=(key, money(received), money(paid), money(received - paid)))

    def draw_spending(self) -> None:
        canvas = self.spending_chart
        canvas.delete("all")
        canvas.create_text(12, 16, text="Top spending categories", anchor="w", fill=_INK, font=("Segoe UI", 10, "bold"))
        if not self.chart_data:
            canvas.create_text(12, 47, text="No spending in this period.", anchor="w", fill=_MUTED, font=("Segoe UI", 10))
            return
        width = max(20, canvas.winfo_width() - 245)
        peak = self.chart_data[0][1]
        for index, (label, cents) in enumerate(self.chart_data):
            y = 42 + 28 * index
            canvas.create_text(12, y, text=label.title(), anchor="w", fill=_INK, font=("Segoe UI", 9))
            canvas.create_rectangle(110, y - 5, 110 + width, y + 5, fill="#E3EEED", outline="")
            canvas.create_rectangle(110, y - 5, 110 + max(2, width * cents / peak), y + 5, fill=_TEAL, outline="")
            canvas.create_text(123 + width, y, text=money(cents), anchor="w", fill=_INK, font=("Segoe UI", 9))

    def reload(self) -> None:
        if self.ledger:
            self.run("Reading transactions…", self.ledger.entries, self._refreshed)

    def _refreshed(self, entries: list[Entry]) -> None:
        self.render(entries)
        self.status.set("Transactions refreshed.")

    def open_report(self) -> None:
        try:
            if not self.report.is_file():
                raise LedgerError("The report is pending. Retry the Excel update in Settings.")
            os.startfile(self.report)
        except (LedgerError, OSError) as exc:
            messagebox.showerror("Open Excel report", _message(exc), parent=self.root)

    def _save(self, items: list[Entry], original: str | None = None) -> None:
        if not self.ledger:
            return

        def done(ids):
            self.undo_ids = ids
            self.undo_button.configure(state="normal")
            if original is not None and self.note.get("1.0", "end-1c") == original:
                self.note.delete("1.0", "end")
            self._changed(f"Saved {len(ids)} {'entry' if len(ids) == 1 else 'entries'}. Undo is available.")

        self.run("Saving transactions…", partial(self.ledger.add, items), done)

    def _changed(self, status: str) -> None:
        def loaded(items):
            self.render(items)
            self.status.set(status)
            self.export()

        def failed(error):
            self.status.set(status + " The list could not refresh.")
            self.excel.set("Excel update pending.")
            self.export()
            messagebox.showwarning("Refresh transactions", _message(error), parent=self.root)

        self.pending = True
        self.run("Refreshing transactions…", self.ledger.entries, loaded, failed)

    def log(self) -> None:
        if not self.ledger or self.busy:
            return
        original = self.note.get("1.0", "end-1c")
        self.log_button.configure(state="disabled")

        def parsed(result):
            self.log_button.configure(state="normal")
            items, warning = result
            known = {(item.kind, item.day, item.cents, item.currency, item.note.strip().casefold()) for item in self.entries}
            if any((item.kind, item.day, item.cents, item.currency, item.note.casefold()) in known for item in items):
                warning = warning or "An entry from the same note already exists. Confirm whether to save another."
            if warning or not self.auto.get():
                items = self.review(items, warning)
            if items:
                self._save(items, original)
            else:
                self.status.set("Draft kept. No entries saved.")

        def failed(error):
            self.log_button.configure(state="normal")
            self.status.set("Draft kept. No entries saved.")
            messagebox.showerror("Could not log note", _message(error), parent=self.root)

        self.run("Reading your note with Ollama…", partial(parse_note, original, self.settings, self.client), parsed, failed)

    def review(self, items: list[Entry], warning: str) -> list[Entry]:
        dialog = tk.Toplevel(self.root)
        dialog.title("Review transactions")
        dialog.transient(self.root)
        dialog.grab_set()
        panel = ttk.Frame(dialog, padding=16)
        panel.pack(fill="both", expand=True)
        ttk.Label(panel, text=warning or "Review the extracted transactions.", wraplength=620).pack(anchor="w", pady=(0, 8))
        tree = self._tree(panel, ("Date", "Type", "Name", "Category", "Payment", "Amount"),
                          (100, 78, 143, 110, 100, 130), min(len(items), 7))
        result: list[Entry] = []

        def redraw():
            tree.delete(*tree.get_children())
            for index, item in enumerate(items):
                tree.insert("", "end", iid=str(index), values=(item.day, item.kind.title(), item.name,
                                                               item.category, item.payment, f"{money(item.cents)} {item.currency}"))

        def change(remove=False):
            selection = tree.selection()
            if not selection:
                return
            index = int(selection[0])
            if remove:
                items.pop(index)
            else:
                item = self.form(items[index], parent=dialog)
                if item:
                    items[index] = item
            redraw()

        def save():
            if not items or any(item.kind == "income" and item.name.casefold() == "n/a" for item in items):
                messagebox.showwarning("Review transactions", "Keep at least one entry and fill in all income sources.", parent=dialog)
                return
            result.extend(items)
            dialog.destroy()

        redraw()
        tree.bind("<Double-1>", lambda _event: change())
        row = ttk.Frame(panel)
        row.pack(fill="x", pady=(8, 0))
        ttk.Button(row, text="Edit selected", command=change).pack(side="left")
        ttk.Button(row, text="Remove selected", command=lambda: change(True)).pack(side="left", padx=4)
        ttk.Button(row, text="Save entries", style="Accent.TButton", command=save).pack(side="right")
        ttk.Button(row, text="Cancel", command=dialog.destroy).pack(side="right", padx=4)
        self.root.wait_window(dialog)
        return result

    def form(self, item: Entry | None = None, *, parent=None) -> Entry | None:
        parent = parent or self.root
        previous_grab = parent.grab_current()
        dialog = tk.Toplevel(parent)
        dialog.title("Edit transaction" if item else "Manual entry")
        dialog.transient(parent)
        dialog.grab_set()
        panel = ttk.Frame(dialog, padding=16)
        panel.pack(fill="both", expand=True)
        fields = (
            ("Type", item.kind if item else "expense", ("expense", "income")),
            ("Date (YYYY-MM-DD)", item.day if item else date.today().isoformat(), ()),
            ("Amount", money(item.cents).replace(",", "") if item else "", ()),
            ("Currency", item.currency if item else self.currency, ()),
            ("Merchant or source", item.name if item else "", ()),
            ("Category", item.category if item and item.category else "OTHERS", tuple(c.value for c in Category)),
            ("Payment method", item.payment if item and item.payment else "Unknown", tuple(p.value for p in PaymentMethod)),
            ("Note", item.note if item else "", ()),
        )
        variables = []
        for row, (label, value, options) in enumerate(fields):
            var = tk.StringVar(value=value)
            variables.append(var)
            ttk.Label(panel, text=label).grid(row=row, column=0, sticky="w", pady=4)
            widget = (ttk.Combobox(panel, textvariable=var, values=options, state="readonly", width=42)
                      if options else ttk.Entry(panel, textvariable=var, width=44))
            widget.grid(row=row, column=1, padx=(12, 0), sticky="ew")
        result: list[Entry] = []

        def save():
            kind, day, amount, code, name, category, payment, note = (var.get().strip() for var in variables)
            try:
                result.append(Entry(kind, day, as_cents(amount), code.upper(), name,
                                    category if kind == "expense" else "", payment if kind == "expense" else "", note))
            except LedgerError as exc:
                messagebox.showerror("Check transaction", str(exc), parent=dialog)
                return
            dialog.destroy()

        ttk.Button(panel, text="Save", style="Accent.TButton", command=save).grid(row=len(fields), column=1, sticky="e", pady=10)
        ttk.Button(panel, text="Cancel", command=dialog.destroy).grid(row=len(fields), column=0, sticky="w", pady=10)
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        self.root.wait_window(dialog)
        if previous_grab and previous_grab.winfo_exists():
            previous_grab.grab_set()
        return result[0] if result else None

    def manual(self) -> None:
        if self.ledger and not self.busy:
            item = self.form()
            if item:
                self._save([item])

    def _selected(self) -> Entry | None:
        selected = self.history.selection()
        if not selected:
            self.status.set("Select a transaction first.")
            return None
        return next((item for item in self.entries if str(item.id) == selected[0]), None)

    def edit(self) -> None:
        if self.busy or not self.ledger:
            return
        previous = self._selected()
        if previous:
            item = self.form(previous)
            if item:
                self.run("Saving changes…", partial(self.ledger.update, previous, item),
                         lambda _value: self._changed("Transaction updated."))

    def delete(self) -> None:
        if self.busy or not self.ledger:
            return
        item = self._selected()
        if item and messagebox.askyesno("Delete transaction", f"Delete {item.name} ({money(item.cents)} {item.currency})?",
                                        parent=self.root):
            def removed(_value):
                if item.id in self.undo_ids:
                    self.undo_ids.remove(item.id)
                    if not self.undo_ids:
                        self.undo_button.configure(state="disabled")
                self._changed("Transaction deleted.")

            self.run("Deleting transaction…", partial(self.ledger.remove, [item.id]), removed)

    def undo(self) -> None:
        if not self.ledger or not self.undo_ids or self.busy:
            return

        def done(_value):
            self.undo_ids = []
            self.undo_button.configure(state="disabled")
            self._changed("Last save undone.")

        self.run("Undoing last save…", partial(self.ledger.remove, self.undo_ids.copy()), done)

    def ask(self) -> None:
        if not self.ledger or self.busy:
            return

        def done(text):
            self.answer.configure(state="normal")
            self.answer.delete("1.0", "end")
            self.answer.insert("1.0", text)
            self.answer.configure(state="disabled")
            self.status.set("Answer ready.")

        question = self.question.get("1.0", "end-1c")
        currency = self.code.get()
        self.run("Working on your question…",
                 lambda: answer_question(question, self.ledger.entries(), currency, self.client),
                 done)

    def check_model(self) -> None:
        def done(names):
            wanted = self.settings.model
            installed = wanted in names or (":" not in wanted and f"{wanted}:latest" in names)
            self.status.set(f"Ollama is ready · {wanted}" if installed else f"Ollama is running. Install {wanted} to enable model features.")

        self.run("Checking local Ollama…", self.client.health, done)

    def choose_import(self) -> None:
        if self.busy:
            return
        chosen = filedialog.askopenfilename(parent=self.root, filetypes=[("Excel workbook", "*.xlsx")])
        if not chosen:
            return
        self.source = Path(chosen)
        if self.source.resolve() == self.report.resolve():
            messagebox.showerror("Import workbook", "Select the original workbook, separate from the generated Excel report.",
                                 parent=self.root)
            return
        if not self.ledger:
            self.start()
            return

        def done(count):
            self._changed(f"Imported {count} transactions." if count else "No new rows imported. Startup import only runs once.")

        self.run("Importing workbook…", partial(self.ledger.import_once, self.source, self.settings.sheet_name, self.currency), done)


def main() -> None:
    settings = Settings()
    parser = argparse.ArgumentParser(description="Local desktop expense tracker")
    parser.add_argument("--data", type=Path, default=Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "Expense Tracker")
    parser.add_argument("--import-from", type=Path, default=settings.xlsx_path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--model", default=os.getenv("EXPENSE_MODEL") or "qwen3.5:4b")
    args = parser.parse_args()
    settings = replace(settings, model=args.model)
    data = args.data.expanduser().resolve()
    source = args.import_from.expanduser().resolve()
    report = (args.report or (Path.home() / "Documents" / "Expense Tracker" / "expenses.xlsx")).expanduser().resolve()
    root = tk.Tk()
    endpoint = urlsplit(settings.ollama_url)
    if endpoint.scheme not in {"http", "https"} or endpoint.hostname not in {"localhost", "127.0.0.1", "::1"}:
        root.withdraw()
        messagebox.showerror("Expense Tracker", "Set OLLAMA_URL to an endpoint on this computer.", parent=root)
        root.destroy()
        return
    if source == report or report.suffix.lower() != ".xlsx":
        root.withdraw()
        messagebox.showerror("Expense Tracker", "Choose a separate .xlsx report path.", parent=root)
        root.destroy()
        return
    try:
        data.mkdir(parents=True, exist_ok=True)
        logging.basicConfig(
            filename=data / "app.log", level=logging.WARNING, encoding="utf-8",
            format="%(asctime)s %(levelname)s %(message)s",
        )
    except OSError as exc:
        root.withdraw()
        messagebox.showerror("Expense Tracker", _message(exc), parent=root)
        root.destroy()
        return
    Desktop(root, settings, data, source, report)
    root.mainloop()


if __name__ == "__main__":
    main()

