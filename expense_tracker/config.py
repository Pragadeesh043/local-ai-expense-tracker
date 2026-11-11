"""Settings. Env vars override defaults. No network config beyond localhost Ollama."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    ollama_url: str = field(default_factory=lambda: os.getenv("OLLAMA_URL", "http://localhost:11434"))
    model: str = field(default_factory=lambda: os.getenv("EXPENSE_MODEL", "llama3.2"))
    xlsx_path: Path = field(
        default_factory=lambda: Path(os.getenv("EXPENSE_XLSX", "expenses.xlsx")).expanduser()
    )
    sheet_name: str = field(default_factory=lambda: os.getenv("EXPENSE_SHEET", "Expenses"))
    default_currency: str = field(default_factory=lambda: os.getenv("EXPENSE_CURRENCY", "INR"))
    sort_order: str = field(default_factory=lambda: os.getenv("EXPENSE_SORT", "desc"))
    monthly_totals: bool = field(
        default_factory=lambda: os.getenv("EXPENSE_TOTALS", "1").strip().lower() not in {"0", "false", "no", "off"}
    )
    timeout_s: float = field(default_factory=lambda: float(os.getenv("EXPENSE_TIMEOUT", "120")))
    retries: int = field(default_factory=lambda: int(os.getenv("EXPENSE_RETRIES", "2")))
