# Local Expense Tracker

Type an expense in plain words. A local Ollama model turns it into a structured row. The row is appended to an `.xlsx` file. Nothing leaves your machine.

## Setup

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
