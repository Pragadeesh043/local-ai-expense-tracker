# Local Expense Tracker

Log expenses and income in plain language using a local Ollama model. The Windows app manages the ledger and refreshes an Excel report. A command-line interface is also available.

## Windows desktop app

Run `python -m expense_tracker.desktop` from the project directory, or launch the packaged `ExpenseTracker.exe`. The desktop app uses an installed copy of Ollama and defaults to `qwen3.5:4b`. Install that model with `ollama pull qwen3.5:4b` if needed. Settings has a **Check Ollama** button. Pass `--model NAME` or set `EXPENSE_MODEL` to use another local model.

- **Log** turns a plain-language note into one or more expenses and income entries. Clear notes can save automatically. Ambiguous notes open an editable review. Undo reverses the last new batch. Manual entry also works with Ollama offline.
- **Dashboard** shows opening balance, income, spending, closing balance, savings rate, merchant and payment breakdowns, and monthly cash flow. The previous closing balance carries into the next month as its opening balance. Use the month and currency selectors to switch views.
- **Transactions** lets you search, edit, and delete recorded entries.
- **Ask** supports calculations and read-only ledger questions. Examples include “15% of 42500,” “How much did I spend on food last month?”, and “What were my five biggest UPI expenses in November?”

The desktop ledger lives in `%LOCALAPPDATA%\Expense Tracker\transactions.sqlite3`. On first use from the project directory, the app imports the existing `expenses.xlsx` once. To select another original workbook, use **Import existing workbook** in Settings or pass `--import-from PATH`. The original workbook remains available. Make transaction corrections in the app.

After each change, the app updates `~/Documents/Expense Tracker/expenses.xlsx`, its generated report. It contains a dashboard, filterable Expense and Income tables, and a monthly summary by currency. If Excel has it locked, transactions are still saved locally. The app retries the report update while running and at the next launch. Give the report its own file path. Use `--report PATH` to choose a different destination or `--data DIRECTORY` for a separate ledger.

To build the standalone Windows executable from a Python environment with the project requirements installed:

```powershell
python -m pip install pyinstaller
python -m PyInstaller --onefile --windowed --name ExpenseTracker launch_desktop.pyw
```

## Command line

The original CLI can still be used with its own workbook.

### Setup

```
pip install -r requirements.txt
ollama pull llama3.2
ollama serve            # if it is not already running
python -m expense_tracker check
```

Python 3.11+ is required.

## Use

```
python -m expense_tracker add "swiggy 450 upi yesterday"
python -m expense_tracker add -y "uber 230 cash last friday"
python -m expense_tracker chat
python -m expense_tracker list -n 5
```

`add` shows the parsed result and asks before saving. Use `-y` to skip the question.

## Settings

Set with environment variables or flags.

| Variable | Default | Flag |
|---|---|---|
| `EXPENSE_MODEL` | `llama3.2` | `--model` |
| `EXPENSE_XLSX` | `expenses.xlsx` | `--file` |
| `EXPENSE_SHEET` | `Expenses` | |
| `EXPENSE_CURRENCY` | `INR` | |
| `OLLAMA_URL` | `http://localhost:11434` | |
| `EXPENSE_TIMEOUT` | `120` (seconds) | |
| `EXPENSE_RETRIES` | `2` | |
| `EXPENSE_SORT` | `desc` (newest on top). Use `asc` for oldest first. | |

To use `phi3`: `EXPENSE_MODEL=phi3 python -m expense_tracker chat`.

## Several expenses in one note

```
python -m expense_tracker add "cab 490 + lunch 330 paid via cash yesterday"
```

This writes two rows. The preview shows both, you confirm once, and both are saved in one write. If the file is locked, neither is saved.

A detail stated once (date, payment method) is applied to every item by the model. Check the preview when you mix payment methods in one note.

## How it works

1. The note goes to Ollama with a JSON schema (from the Pydantic model) in the `format` field. The model returns `{"expenses": [...]}`, one object per item. A bare list or a single object is also accepted.
2. The model only copies the date phrase ("last Friday"). Python resolves it to a real date. Small models are weak at date math.
3. A date phrase from the model is used only if it really appears in your note. Otherwise Python scans the note for "yesterday", "N days ago", "last friday", "3 nov" and similar. If nothing is found, the date is today.
4. Pydantic validates and normalizes every field. Bad output is retried.
5. `openpyxl` appends the rows. Headers and styling are created once and kept.

The schema is in `expense_tracker/schema.py`. Edit the `Category` enum there to change categories.

## CLI monthly totals and dashboard

~~~powershell
python -m expense_tracker summarize
~~~

The CLI rebuilds its workbook's Dashboard after each addition and when you run summarize. The Dashboard shows lifetime income, spending, and net balance, with a category table and pie chart for each month that has spending in the configured reporting currency. Excel or LibreOffice calculates the workbook formulas when you open it.

The Expenses sheet has monthly TOTAL rows. A month with multiple currencies gets one total per currency. Set EXPENSE_TOTALS=0 to omit these rows. The Dashboard is regenerated on each write.

## Sorting and duplicates

- Every add re-sorts all data rows by date. Rows with the same date keep their order. Rows with no readable date go last.
- The header, freeze pane and column widths are not touched. Each row keeps its own cell styles when it moves. Formulas in data rows are re-pointed to their new row. Keep totals above the table or on another sheet, because a total row under the data would be sorted too.
- Blank rows between data rows are removed when the sheet is rewritten.
- If the model returns the same item twice, the extra copy is dropped. Identical items are kept only as many times as their amount appears in your note, so "coffee 50 + tea 50" still gives two rows.
- Rows already in your file are never removed, so delete earlier duplicates by hand.

## Notes

- Close the `.xlsx` in Excel before saving. Excel locks the file.
- Dates with no year that fall in the future are read as last year.
- Numeric dates like `3/11` are read day first.
- GridCraft: this tool writes standard `.xlsx`, which GridCraft opens. I did not verify a GridCraft API or CLI, so none is used.

## Tests

```
python -m pytest -q
```
