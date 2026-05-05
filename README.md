# Trials Queue Dashboard

This is a **local desktop UI**, not a hosted website. Each user runs the app on their computer and the app reads/writes the shared `testing QUEUE.xlsx` workbook directly. That keeps Excel/SharePoint as the backend/source of truth while giving the two teams different screens.

## What it does

- Uses the Excel workbook as the backend.
- Gives Pre-prep a local form for adding line items while still allowing the team to edit the workbook directly in SharePoint.
- Lets Pre-prep mark `Ready for Prep` with `X` so coordinators can see the row.
- Gives Device Coordinators a filtered dashboard showing only prep-ready or 2nd-check-ready rows.
- Groups coordinator work into priority lanes: Expedites, Funded Rentals, Accessories, Requested Ship Dates, and Regular Daily Queue.
- Maintains separate FIFO queues for prep and 2nd-check work in a small JSON state file.
- Offers work to the first coordinator in the queue for 2 minutes by default.
- Moves a coordinator to the bottom of that FIFO queue when they skip, complete, or let an offer expire.
- Writes claim, completion, assignment, and QA-ready updates back into the Excel workbook.

## Run the local app

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

On Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

## Point it at the SharePoint workbook

If the workbook is synced locally from SharePoint, point the app at that file:

```bash
QUEUE_EXCEL_PATH="/path/to/SharePoint/testing QUEUE.xlsx" python app.py
```

For multiple coordinators to share FIFO queue membership, also put the queue state JSON somewhere synced/shared:

```bash
QUEUE_STATE_PATH="/path/to/SharePoint/queue_state.json" python app.py
```

Optional environment variables:

- `QUEUE_EXCEL_PATH`: path to the Excel workbook. Defaults to this repo's `testing QUEUE.xlsx`.
- `QUEUE_STATE_PATH`: JSON file used for FIFO queue membership. Defaults to `data/queue_state.json`.
- `CLAIM_TIMEOUT_SECONDS`: offer timeout. Defaults to `120`.
- `QUEUE_POLL_SECONDS`: how often the local app refreshes and checks for offers. Defaults to `10`.

## Workbook expectations

The included workbook already has the expected headers. The app recognizes common header aliases, including `ReadyForPrep` / `Ready for Prep`; it also falls back to column H if a team workbook marks readiness there. Mark ready cells with `X`, `Yes`, `Y`, `True`, or `1`.

## Important SharePoint note

Because this app writes directly to Excel, avoid having the workbook open in desktop Excel while coordinators are claiming/completing work. Excel file locks can block writes. A synced SharePoint file is still the backend, but this is not a database server.
