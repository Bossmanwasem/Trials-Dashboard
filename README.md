# Daily Queue Dashboard

A Windows desktop MVP for the internal Daily Queue workflow. Pre-prep users can keep maintaining the existing RRQ and Daily Queue Excel workbooks, while prep and QA users work from this app instead of directly editing the Daily Queue workbook.

## What this MVP does

- Imports queue rows from an existing Daily Queue Excel workbook (`.xlsx` / `.xlsm`).
- Stores imported jobs and lane state in a SQLite database (`daily_queue.sqlite3` by default).
- Saves each workstation user's name and initials locally in a JSON profile file, not in the shared database.
- Shows a one-screen CustomTkinter dashboard with workflow tabs/lanes:
  - **Ready to Claim**
  - **Currently Being Prepped**
  - **Ready for QA**
  - **Currently Being QA**
  - **Completed**
- Lets users claim prep with their saved initials automatically.
- Uses atomic SQLite transactions so two app users cannot claim the same prep or QA job at the same time.
- Moves jobs between lanes from SQLite status, not Excel formulas.
- Records a status history entry whenever a job moves lanes.
- Writes only the mapped Excel cells for workflow updates:
  - Prep claim writes the user's initials to column `I`.
  - QA completion writes `INITIALS-Yes` to column `J`.

## Lane behavior

| Lane | What appears here | Available action |
| --- | --- | --- |
| Ready to Claim | Jobs where Excel column `H` is `X`, Excel column `I` is blank, and SQLite status is `Ready Prep`. | `Claim for Prep` |
| Currently Being Prepped | Jobs claimed for prep with SQLite status `In Prep`. | `Finish Prep`, shown only to the user who claimed the job. |
| Ready for QA | Jobs where prep is finished and QA is not complete, with SQLite status `Ready QA`. | `Claim QA` |
| Currently Being QA | Jobs claimed for QA with SQLite status `In QA`. | `Finish QA`, shown only to the user who claimed QA. |
| Completed | Jobs where QA is complete with SQLite status `Completed`. | No workflow action. |

Expedite sections are sorted to the top of each lane.

## Project layout

| File | Purpose |
| --- | --- |
| `app.py` | CustomTkinter UI, profile dialog, workflow lane rendering, and button actions. |
| `db.py` | SQLite schema, imports, claims, lane transitions, and status history. |
| `excel_sync.py` | Excel import/export helpers and column mapping config. |
| `user_profile.py` | Local workstation profile JSON persistence and validation. |
| `models.py` | Shared status and lane constants. |
| `app_config.py` | App name, default database path, and default local profile path. |
| `tests/test_db.py` | pytest coverage for lane imports, claims, transitions, and history. |
| `tests/test_user_profile.py` | pytest coverage for local profile persistence. |
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

On first launch, the app asks for your name and initials. The profile is saved locally on that computer:

- Windows: `%LOCALAPPDATA%\DailyQueueDashboard\profile.json`
- Other platforms: `~/.daily_queue_dashboard/profile.json`

Use **Settings/Profile** in the top bar to update the saved name or initials later.

Then:

1. Click **Import from Excel** and select the Daily Queue workbook.
2. Open the **Ready to Claim** lane.
3. Click **Claim for Prep** on a job. The app uses the active user's saved initials automatically and writes them to Excel column `I`.
4. The job immediately leaves **Ready to Claim** and appears in **Currently Being Prepped**.
5. The claiming user clicks **Finish Prep** to move the job to **Ready for QA**. This is app-only lane state and does not write a separate Excel field.
6. A QA user clicks **Claim QA** to move the job to **Currently Being QA**.
7. The QA user who claimed the job clicks **Finish QA** to move it to **Completed** and write `INITIALS-Yes` to Excel column `J`.

## Shared database option

By default, the app creates `daily_queue.sqlite3` beside the app. For multiple users, place the app and database on a shared network location that all users can access. SQLite write transactions are intentionally short and use atomic claim updates to prevent double-claiming.

The user profile file should stay local on each computer and should not be placed in the shared SQLite database.

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
    "ready_for_prep": "H",
    "prep_initials": "I",
    "qa_done": "J",
}
PREP_USERNAME_COLUMN = "I"
QA_DONE_COLUMN = "J"
QA_DONE_SUFFIX = "-Yes"
```

Only `PREP_USERNAME_COLUMN` and `QA_DONE_COLUMN` are written by the app. The rest are read into SQLite for dashboard display and initial lane placement.

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
4. Leave each user's profile in their local workstation profile path.
5. Ask users to close Excel before clicking app actions that write columns `I` or `J`; Excel file locks can prevent saving.

## Run tests

```powershell
pytest
```

The current tests focus on SQLite workflow safety and local profile persistence and do not require opening the UI.

## Notes and limitations

- This is an MVP with one dashboard screen and tabbed lanes.
- The app saves targeted Excel cell changes in place using `openpyxl`; it does not intentionally rewrite unrelated cells.
- If an Excel save fails because the workbook is open or locked, the app keeps the SQLite workflow state and shows a warning.
- Prep completion and QA claiming are app-only lane transitions; Excel still receives only column `I` prep initials and column `J` QA completion.
