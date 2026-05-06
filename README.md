# Daily Queue Dashboard

A Windows desktop MVP for the internal Daily Queue workflow. Pre-prep users can keep maintaining the existing RRQ and Daily Queue Excel workbooks, while prep and QA users work from this app instead of directly editing the Daily Queue workbook.

## What this MVP does

- Imports queue rows from an existing Daily Queue Excel workbook (`.xlsx` / `.xlsm`).
- Stores imported jobs in a SQLite database (`daily_queue.sqlite3` by default).
- Shows a one-screen CustomTkinter dashboard grouped by the same dashed section rows used in the workbook.
- Lets users enter initials and claim a selected job for prep.
- Uses an atomic SQLite update so two app users cannot claim the same job at the same time.
- Marks QA done for a selected job.
- Writes only the mapped Excel cells for workflow updates:
  - Prep claim writes the user's initials to column `I`.
  - QA completion writes `INITIALS-Yes` to column `J`.
- Keeps visual status in the app so users can quickly see imported, prep-claimed/ready-for-QA, and QA-complete rows.

## Project layout

| File | Purpose |
| --- | --- |
| `app.py` | CustomTkinter desktop UI and button workflow. |
| `db.py` | SQLite schema, imports, claims, and QA completion logic. |
| `excel_sync.py` | Excel import/export helpers and column mapping config. |
| `models.py` | Shared status constants. |
| `app_config.py` | App name and default database path. |
| `tests/test_db.py` | pytest coverage for import preservation, single-winner claiming, and QA completion. |
| `testing QUEUE.xlsx` | Example workbook included in this repository. |

## Install for development

> Recommended Python version: Python 3.11 or newer.

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

If your workstation has multiple Python versions, `py -3.11` can be replaced with the exact installed version, such as `py -3.12`.

## Run the app

```powershell
.\.venv\Scripts\Activate.ps1
python app.py
```

Then:

1. Enter your initials in the top bar.
2. Click **Import from Excel** and select the Daily Queue workbook.
3. Click a queue row to select it.
4. Click **Claim for Prep** to claim the row and write your initials to Excel column `I`.
5. Click **Mark QA Done** to mark QA complete and write `INITIALS-Yes` to Excel column `J`.

## Shared database option

By default, the app creates `daily_queue.sqlite3` beside the app. For multiple users, place the app and database on a shared network location that all users can access. SQLite write transactions are intentionally short and use an atomic claim update to prevent double-claiming.

For a future production rollout, consider adding a small settings screen or shortcut argument for choosing a shared database path.

## Excel mapping

The workbook mapping lives at the top of `excel_sync.py`:

```python
QUEUE_SHEET_NAME = "Daily Queue"
HEADER_ROW = 1
FIRST_DATA_ROW = 2
COLUMN_MAP = {
    "last_name": "A",
    "first_name": "B",
    "device": "C",
    "loan_type": "D",
    "queue_date": "E",
    "vocabulary": "F",
    "notes": "G",
}
PREP_USERNAME_COLUMN = "I"
QA_DONE_COLUMN = "J"
QA_DONE_SUFFIX = "-Yes"
```

Only `PREP_USERNAME_COLUMN` and `QA_DONE_COLUMN` are written by the app. The rest are read into SQLite for dashboard display.

## Build a Windows `.exe`

Install dependencies first, then run PyInstaller from an activated virtual environment:

```powershell
python -m PyInstaller --noconfirm --onefile --windowed --name DailyQueueDashboard app.py
```

The executable will be created at:

```text
dist\DailyQueueDashboard.exe
```

Suggested deployment steps:

1. Copy `dist\DailyQueueDashboard.exe` to a shared folder or each user's workstation.
2. Keep the Daily Queue workbook in its normal shared location.
3. Put `daily_queue.sqlite3` in a shared folder if everyone should see the same app status.
4. Ask users to close Excel before clicking app actions that write columns `I` or `J`; Excel file locks can prevent saving.

## Run tests

```powershell
pytest
```

The current tests focus on SQLite workflow safety and do not require opening the UI.

## Notes and limitations

- This is an MVP with one dashboard screen.
- The app saves targeted Excel cell changes in place using `openpyxl`; it does not intentionally rewrite unrelated cells.
- If an Excel save fails because the workbook is open or locked, the app keeps the SQLite workflow state and shows a warning.
- The current QA button does not require a separate QA claim step; it marks QA complete directly for the selected job.
