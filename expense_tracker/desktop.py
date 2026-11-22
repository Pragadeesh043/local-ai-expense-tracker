"""Windows-friendly desktop interface. Run with python -m expense_tracker.desktop."""
from __future__ import annotations

import argparse
import ctypes
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
from .ledger_report import breakdown, month_groups, monthly_balances, sync_report
from .llm import OllamaClient, OllamaError
from .schema import Category, PaymentMethod


_BG = "#EFF3F8"
_INK = "#172C44"
_TEAL = "#08796D"
_MUTED = "#53677F"
_WHITE = "#FFFFFF"
_BORDER = "#D9E3ED"
_NAV = "#132A43"
_NAV_ACTIVE = "#25435D"
_NAV_TEXT = "#B9CBDE"
_MINT = "#72E4CE"
_GREEN = "#08796D"
_RED = "#AA4050"
_BAR_COLORS = ("#08796D", "#3265A3", "#7B55A8", "#A56219")


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
        self.motion = tk.BooleanVar(value=True)
        self._chart_job: str | None = None
        self._chart_fraction = 1.0
        self._save_job: str | None = None
        root.title("Expense Tracker")
        root.geometry("1140x760")
        root.minsize(760, 640)
        root.configure(bg=_BG)
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.report_callback_exception = self.callback_error
        self._layout()
        root.after_idle(self.note.focus_set)
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
        style.configure("Section.TLabel", font=("Segoe UI", 25, "bold"), foreground=_INK)
        style.configure("Subheading.TLabel", font=("Segoe UI", 14, "bold"), foreground=_INK)
        style.configure("Muted.TLabel", foreground=_MUTED)
        style.configure("Surface.TFrame", background=_WHITE)
        style.configure("Surface.TLabel", background=_WHITE, foreground=_INK)
        style.configure("SurfaceMuted.TLabel", background=_WHITE, foreground=_MUTED)
        style.configure("SurfaceHeading.TLabel", background=_WHITE, font=("Segoe UI", 13, "bold"))
        style.configure("Surface.TCheckbutton", background=_WHITE)
        style.map("Surface.TCheckbutton", background=[("active", "#E7F4F1")])
        style.configure("TButton", background=_WHITE, bordercolor=_BORDER, lightcolor=_WHITE, darkcolor=_BORDER,
                        padding=(10, 7))
        style.map("TButton", background=[("active", "#E3ECF6")], foreground=[("disabled", "#65768A")])
        style.configure("Accent.TButton", background=_TEAL, foreground=_WHITE, bordercolor=_TEAL,
                        lightcolor=_TEAL, darkcolor=_TEAL, font=("Segoe UI", 10, "bold"), padding=(14, 8))
        style.map("Accent.TButton", background=[("disabled", "#C5DBD7"), ("active", "#05675E")],
                  foreground=[("disabled", "#345C56"), ("!disabled", _WHITE)])
        style.configure("TNotebook", borderwidth=0, background=_BG, bordercolor=_BORDER,
                        lightcolor=_BORDER, darkcolor=_BORDER)
        style.configure("TNotebook.Tab", padding=(14, 8))
        style.map("TNotebook.Tab", background=[("selected", _WHITE), ("active", "#DCE8F3")])
        style.configure("Pages.TNotebook", padding=0, borderwidth=0, bordercolor=_BG,
                        lightcolor=_BG, darkcolor=_BG)
        style.layout("Pages.TNotebook.Tab", [])
        style.configure("Treeview", rowheight=31, fieldbackground=_WHITE, background=_WHITE,
                        borderwidth=0, bordercolor=_BORDER, lightcolor=_BORDER, darkcolor=_BORDER,
                        font=("Segoe UI", 10))
        style.configure("Treeview.Heading", background="#E5EDF6", relief="flat", foreground=_INK,
                        font=("Segoe UI", 10, "bold"), padding=(5, 7))
        style.map("Treeview", background=[("selected", "#D4E8F8")], foreground=[("selected", _INK)])
        style.map("Treeview.Heading", background=[("active", "#D8E7F4")])
        style.configure("Horizontal.TProgressbar", troughcolor=_BORDER, background=_TEAL,
                        lightcolor=_TEAL, darkcolor=_TEAL, borderwidth=0, thickness=3)
        style.configure("TEntry", fieldbackground=_WHITE, padding=6)
        style.configure("TCombobox", padding=4)

        rail = tk.Frame(self.root, bg=_NAV, width=172)
        rail.pack(side="left", fill="y")
        rail.pack_propagate(False)
        brand = tk.Frame(rail, bg=_NAV, padx=17, pady=21)
        brand.pack(fill="x")
        tk.Frame(brand, width=27, height=4, bg=_MINT).pack(anchor="w", pady=(0, 8))
        tk.Label(brand, text="Expense\nTracker", justify="left", bg=_NAV, fg=_WHITE, borderwidth=0,
                 font=("Segoe UI", 19, "bold")).pack(anchor="w")
        tk.Button(rail, text="Open Excel report", command=self.open_report, anchor="w",
                  bg=_NAV, fg=_NAV_TEXT, activebackground=_NAV_ACTIVE, activeforeground=_WHITE,
                  relief="flat", borderwidth=0, padx=15, pady=10, cursor="hand2").pack(
                      side="bottom", fill="x", padx=8, pady=(7, 15))

        shell = ttk.Frame(self.root, padding=(22, 14, 22, 10))
        shell.pack(side="left", fill="both", expand=True)
        self.tabs = ttk.Notebook(shell, style="Pages.TNotebook", takefocus=False)
        self.tabs.pack(fill="both", expand=True)
        self.nav_buttons: list[tk.Button] = []
        nav = tk.Frame(rail, bg=_NAV)
        nav.pack(fill="x", padx=8, pady=(7, 0))
        for index, (name, build) in enumerate((
            ("Log", self._log_tab), ("Dashboard", self._dash_tab),
            ("Transactions", self._history_tab), ("Ask", self._ask_tab), ("Settings", self._settings_tab),
        )):
            page = ttk.Frame(self.tabs, padding=(1, 8, 1, 8))
            self.tabs.add(page, text=name)
            build(page)
            button = tk.Button(
                nav, text=name, command=partial(self.show_page, index), anchor="w", bg=_NAV, fg=_NAV_TEXT,
                activebackground=_NAV_ACTIVE, activeforeground=_WHITE, relief="flat", borderwidth=0,
                padx=16, pady=10, cursor="hand2", font=("Segoe UI", 10, "bold"),
                highlightthickness=1, highlightbackground=_NAV, highlightcolor=_MINT,
            )
            button.pack(fill="x", pady=2)
            self.nav_buttons.append(button)
        self.tabs.bind("<<NotebookTabChanged>>", self._page_changed)
        self._page_changed()

        footer = ttk.Frame(shell)
        footer.pack(side="bottom", fill="x", before=self.tabs, pady=(4, 0))
        track = ttk.Frame(footer, height=4)
        track.pack(fill="x", pady=(0, 8))
        track.pack_propagate(False)
        self.progress = ttk.Progressbar(track, mode="indeterminate")
        self.progress.pack(fill="both", expand=True)
        status_row = ttk.Frame(footer)
        status_row.pack(fill="x")
        status_row.columnconfigure(0, weight=1)
        status_row.columnconfigure(1, weight=1)
        ttk.Label(status_row, textvariable=self.status, style="Muted.TLabel", wraplength=345).grid(row=0, column=0, sticky="w")
        ttk.Label(status_row, textvariable=self.excel, style="Muted.TLabel", wraplength=300, justify="right").grid(
            row=0, column=1, sticky="e", padx=(12, 0))

    def show_page(self, index: int) -> None:
        self.tabs.select(index)
        if index == 0:
            self.note.focus_set()
        elif index == 3:
            self.question.focus_set()

    def _page_changed(self, _event=None) -> None:
        selected = self.tabs.index("current")
        for index, button in enumerate(self.nav_buttons):
            active = index == selected
            button.configure(bg=_NAV_ACTIVE if active else _NAV, fg=_WHITE if active else _NAV_TEXT,
                             highlightbackground=_NAV_ACTIVE if active else _NAV)
        self.animate_spending()

    def _animate_loading(self, running: bool) -> None:
        self.progress.stop()
        if running and self.motion.get():
            self.progress.configure(mode="indeterminate")
            self.progress.start(15)
        else:
            self.progress.configure(value=45 if running else 0, mode="determinate")

    def _motion_changed(self) -> None:
        self._animate_loading(self.busy)
        self.animate_spending()

    @staticmethod
    def _surface(parent, *, padding=(15, 12)) -> ttk.Frame:
        border = tk.Frame(parent, bg=_BORDER, padx=1, pady=1)
        border.pack(fill="x", pady=(0, 12))
        panel = ttk.Frame(border, style="Surface.TFrame", padding=padding)
        panel.pack(fill="both", expand=True)
        return panel

    @staticmethod
    def _text(parent, height: int, *, readonly: bool = False) -> tk.Text:
        widget = tk.Text(parent, height=height, font=("Segoe UI", 11), relief="flat", bd=0, highlightthickness=1, highlightbackground=_BORDER,
                         highlightcolor=_TEAL, padx=12, pady=10, wrap="word", bg=_WHITE, fg=_INK,
                         insertbackground=_INK, selectbackground="#D4E8F8", selectforeground=_INK, undo=True)
        if readonly:
            widget.configure(state="disabled")
        return widget

    def _log_tab(self, page) -> None:
        ttk.Label(page, text="Log transactions", style="Section.TLabel").pack(anchor="w", pady=(0, 13))
        composer = self._surface(page)
        ttk.Label(composer, text="Transaction note", style="SurfaceHeading.TLabel").pack(anchor="w")
        ttk.Label(composer, text="Example: salary 75000 today + cab 490 UPI yesterday",
                  style="SurfaceMuted.TLabel", wraplength=460).pack(anchor="w", pady=(3, 9))
        self.note = self._text(composer, 4)
        self.note.pack(fill="x")
        self.note.bind("<Control-Return>", lambda _event: (self.log(), "break")[1])
        actions = ttk.Frame(composer, style="Surface.TFrame")
        actions.pack(fill="x", pady=(10, 8))
        self.log_button = ttk.Button(actions, text="Log note", style="Accent.TButton", command=self.log)
        self.log_button.pack(side="left")
        ttk.Button(actions, text="Manual entry", command=self.manual).pack(side="left", padx=6)
        self.undo_button = ttk.Button(actions, text="Undo last save", command=self.undo, state="disabled")
        self.undo_button.pack(side="left")
        ttk.Label(actions, text="Ctrl+Enter", style="SurfaceMuted.TLabel").pack(side="right", padx=3)
        ttk.Checkbutton(composer, text="Automatically save verified notes", variable=self.auto,
                        style="Surface.TCheckbutton").pack(anchor="w")
        ttk.Label(page, text="Recent transactions", style="Subheading.TLabel").pack(anchor="w", pady=(1, 6))
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
        cards.pack(fill="x", pady=(3, 13))
        labels = ("Opening balance", "Income", "Spending", "Closing balance", "Savings rate")
        self.cards = [tk.StringVar() for _ in labels]
        self.card_values: list[tk.Label] = []
        panels: list[tuple[tk.Frame, tk.Label]] = []
        accents = ("#58789B", _GREEN, _RED, _MINT, "#68499B")
        for i, (label, value, accent) in enumerate(zip(labels, self.cards, accents, strict=True)):
            primary = label == "Closing balance"
            background = _NAV if primary else _WHITE
            cards.columnconfigure(i, weight=1, uniform="cards")
            frame = tk.Frame(cards, bg=background, highlightbackground=_NAV if primary else _BORDER, highlightthickness=1)
            frame.grid(row=0, column=i, sticky="nsew", padx=(0, 7) if i < len(labels) - 1 else 0)
            tk.Frame(frame, height=3, bg=accent).pack(fill="x")
            caption = tk.Label(frame, text=label, bg=background, fg=_NAV_TEXT if primary else _MUTED,
                               anchor="w", font=("Segoe UI", 9, "bold"))
            caption.pack(fill="x", padx=9, pady=(8, 1))
            panels.append((frame, caption))
            figure = tk.Label(frame, textvariable=value, bg=background, fg=_WHITE if primary else _INK,
                              anchor="w", font=("Segoe UI", 14, "bold"))
            figure.pack(fill="x", padx=9, pady=(0, 10))
            self.card_values.append(figure)

        current_columns = len(panels)

        def arrange_cards(event) -> None:
            nonlocal current_columns
            columns = 3 if event.width < 760 else len(panels)
            if columns == current_columns:
                return
            current_columns = columns
            compact = columns == 3
            self.spending_chart.configure(height=99 if compact else 143)
            for index, (frame, caption) in enumerate(panels):
                row, column = divmod(index, columns)
                cards.columnconfigure(index, weight=int(index < columns), uniform="cards" if index < columns else "")
                frame.grid_configure(row=row, column=column, pady=(0, 5) if compact else 0,
                                     padx=0 if column == columns - 1 else (0, 7))
                caption.pack_configure(pady=(4 if compact else 8, 1))
                figure = self.card_values[index]
                figure.configure(font=("Segoe UI", 12 if compact else 14, "bold"))
                figure.pack_configure(pady=(0, 5 if compact else 10))

        cards.bind("<Configure>", arrange_cards)
        self.chart_data: list[tuple[str, int]] = []
        chart = self._surface(page, padding=(15, 7))
        self.spending_chart = tk.Canvas(chart, height=143, bg=_WHITE, highlightthickness=0)
        self.spending_chart.pack(fill="x")
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
        self.trend = self._tree(trend, ("Month", "Opening balance", "Income", "Spending", "Closing balance"),
                                (110, 125, 120, 120, 125), 7)

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
        ttk.Label(page, text="Ask about your money", style="Section.TLabel").pack(anchor="w", pady=(0, 13))
        composer = self._surface(page)
        ttk.Label(composer, text="Your question", style="SurfaceHeading.TLabel").pack(anchor="w")
        ttk.Label(composer, text="Try “How much did I spend on food last month?” or “15% of 42500”.",
                  style="SurfaceMuted.TLabel", wraplength=485).pack(anchor="w", pady=(3, 9))
        self.question = self._text(composer, 3)
        self.question.pack(fill="x")
        self.question.bind("<Control-Return>", lambda _event: (self.ask(), "break")[1])
        ttk.Button(composer, text="Ask", style="Accent.TButton", command=self.ask).pack(anchor="w", pady=(9, 0))
        ttk.Label(page, text="Answer", style="Subheading.TLabel").pack(anchor="w", pady=(0, 6))
        self.answer = self._text(page, 11, readonly=True)
        self.answer.pack(fill="both", expand=True)

    def _settings_tab(self, page) -> None:
        ttk.Label(page, text="Settings", style="Section.TLabel").pack(anchor="w", pady=(0, 13))
        connection = self._surface(page)
        ttk.Label(connection, text="Local setup", style="SurfaceHeading.TLabel").pack(anchor="w", pady=(0, 5))
        for label, value in (
            ("Ollama model", self.settings.model), ("Ollama endpoint", self.settings.ollama_url),
            ("Data folder", str(self.data)), ("Excel report", str(self.report)),
        ):
            row = ttk.Frame(connection, style="Surface.TFrame")
            row.pack(fill="x", pady=5)
            ttk.Label(row, text=label, width=16, style="SurfaceMuted.TLabel").pack(side="left", anchor="n")
            ttk.Label(row, text=value, style="Surface.TLabel", wraplength=375).pack(side="left", fill="x")
        row = ttk.Frame(connection, style="Surface.TFrame")
        row.pack(fill="x", pady=(9, 0))
        ttk.Button(row, text="Check Ollama", command=self.check_model).pack(side="left")
        ttk.Button(row, text="Import existing workbook", command=self.choose_import).pack(side="left", padx=8)
        ttk.Button(row, text="Retry Excel update", command=lambda: self.export(force=True)).pack(side="left")
        ttk.Checkbutton(page, text="Animate charts and activity", variable=self.motion,
                        command=self._motion_changed).pack(anchor="w", pady=4)

    @staticmethod
    def _tree(parent, columns, widths, height) -> ttk.Treeview:
        body = ttk.Frame(parent)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=columns, show="headings", height=height, selectmode="browse")
        scroll = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        for col, width in zip(columns, widths, strict=True):
            anchor = "e" if col in {
                "Amount", "Spent", "Share", "Income", "Spending", "Opening balance", "Closing balance",
            } else "w"
            tree.heading(col, text=col, anchor=anchor)
            tree.column(col, width=width, minwidth=52, anchor=anchor)
        tree.tag_configure("alternate", background="#F5F8FC")
        tree.tag_configure("saved", background="#DFF5EB")
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
        self._animate_loading(True)
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
        self._animate_loading(False)
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
        for index, item in enumerate(entries[:7]):
            self.recent.insert("", "end", iid=str(item.id), tags=("alternate",) if index % 2 else (),
                               values=(item.day, item.kind.title(), item.name, f"{money(item.cents)} {item.currency}"))

    def refresh_history(self) -> None:
        query = self.filter.get().strip().casefold()
        self.history.delete(*self.history.get_children())
        for index, item in enumerate(self.entries):
            searchable = " ".join((item.day, item.kind, item.name, item.category, item.payment, item.note, item.currency))
            if query in searchable.casefold():
                self.history.insert("", "end", iid=str(item.id), tags=("alternate",) if index % 2 else (), values=(
                    item.day, item.kind.title(), item.name, item.category, item.payment,
                    f"{money(item.cents)} {item.currency}",
                ))

    def refresh_dashboard(self, _event=None) -> None:
        months = {date.today().isoformat()[:7], *(item.day[:7] for item in self.entries)}
        self.months["values"] = sorted(months, reverse=True)
        self.codes["values"] = sorted({self.currency, *(item.currency for item in self.entries)})
        currency = self.code.get()
        month = self.month.get()
        groups = month_groups([item for item in self.entries if item.currency == currency])
        rows = groups.setdefault((month, currency), [])
        balances = monthly_balances(groups)
        balance = balances[(month, currency)]
        values = [f"{money(cents)} {currency}" for cents in (balance.opening, balance.income, balance.spending, balance.closing)]
        values.append(f"{balance.net_change / balance.income:.1%}" if balance.income else "—")
        for var, value, widget, accent in zip(self.cards, values, self.card_values,
                                              (_INK, _GREEN, _RED, _WHITE, "#68499B"), strict=True):
            var.set(value)
            negative = value.startswith("-")
            widget.configure(fg=("#FFC2CA" if accent == _WHITE else _RED) if negative else accent)
        for field, tree in self.breakdowns.items():
            tree.delete(*tree.get_children())
            for label, cents in breakdown(rows, field):
                tree.insert("", "end", values=(label, money(cents), f"{cents / balance.spending:.1%}"))
        self.chart_data = breakdown(rows, "category")[:4]
        self.animate_spending()
        self.trend.delete(*self.trend.get_children())
        for (key, _), value in reversed(balances.items()):
            self.trend.insert("", "end", values=(key, money(value.opening), money(value.income),
                                                 money(value.spending), money(value.closing)))

    def animate_spending(self) -> None:
        if self._chart_job:
            self.root.after_cancel(self._chart_job)
            self._chart_job = None
        if not self.motion.get() or not self.chart_data or self.tabs.index("current") != 1:
            self._chart_fraction = 1.0
            self.draw_spending()
            return

        def frame(step: int) -> None:
            self._chart_job = None
            self._chart_fraction = 1 - (1 - step / 12) ** 3
            self.draw_spending()
            if step < 12:
                self._chart_job = self.root.after(16, frame, step + 1)

        frame(0)

    def draw_spending(self) -> None:
        canvas = self.spending_chart
        canvas.delete("all")
        compact = int(canvas.cget("height")) < 110
        canvas.create_text(0, 13 if compact else 16, text="Top spending categories", anchor="w", fill=_INK,
                           font=("Segoe UI", 10 if compact else 12, "bold"))
        if not self.chart_data:
            canvas.create_text(0, 51, text="No spending in this period.", anchor="w", fill=_MUTED, font=("Segoe UI", 10))
            return
        width = canvas.winfo_width()
        start, finish = 121, max(141, width - 113)
        peak = self.chart_data[0][1]
        for index, ((label, cents), color) in enumerate(zip(self.chart_data, _BAR_COLORS)):
            y = 35 + 18 * index if compact else 47 + 26 * index
            canvas.create_text(0, y, text=label.title(), anchor="w", fill=_INK, font=("Segoe UI", 9, "bold"))
            canvas.create_line(start, y, finish, y, fill="#E6EDF6", width=8, capstyle="round")
            length = (finish - start) * cents / peak * self._chart_fraction
            if length:
                canvas.create_line(start, y, start + length, y, fill=color, width=8, capstyle="round")
            canvas.create_text(width - 1, y, text=money(cents), anchor="e", fill=_INK, font=("Segoe UI", 9))

    def highlight_saved(self, ids: list[int]) -> None:
        if self._save_job:
            self.root.after_cancel(self._save_job)
        selected = set(map(str, ids))
        rows = self.recent.get_children()
        for index, row in enumerate(rows):
            tags = ("saved",) if self.motion.get() and row in selected else ("alternate",) if index % 2 else ()
            self.recent.item(row, tags=tags)

        def reset() -> None:
            self._save_job = None
            for index, row in enumerate(self.recent.get_children()):
                self.recent.item(row, tags=("alternate",) if index % 2 else ())

        self._save_job = self.root.after(1350, reset) if self.motion.get() else None

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
            self._changed(f"Saved {len(ids)} {'entry' if len(ids) == 1 else 'entries'}. Undo is available.", saved_ids=ids)

        self.run("Saving transactions…", partial(self.ledger.add, items), done)

    def _changed(self, status: str, *, saved_ids: list[int] | None = None) -> None:
        def loaded(items):
            self.render(items)
            if saved_ids:
                self.highlight_saved(saved_ids)
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
    if os.name == "nt":
        ctypes.windll.user32.SetProcessDPIAware()
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

