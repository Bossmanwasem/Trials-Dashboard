# Trials Queue Dashboard

A lightweight Flask dashboard wrapped around the shared `testing QUEUE.xlsx` workbook. Pre-prep users can continue working from the Excel file in SharePoint, while Device Coordinators get a filtered web dashboard for ready devices, priority lanes, prep claiming, and 2nd-check claiming.

## What it does

- Reads and writes the Excel queue file directly.
- Shows Pre-prep a friendly add-row form plus a workbook preview.
- Shows Device Coordinators only rows that are ready for prep or ready for 2nd check.
- Groups visible work into lanes: Expedites, Funded Rentals, Accessories, Requested Ship Dates, and Regular Daily Queue.
- Maintains separate FIFO queues for prep and 2nd checks.
- Offers an order to the first coordinator in queue for 2 minutes by default.
- Moves a coordinator to the bottom of the queue when they skip, complete, or let an offer expire.
- Updates claim, completion, assignment, and QA-ready fields back into Excel.

## Workbook expectations

The included workbook already has the expected headers. The app recognizes common header aliases, including `ReadyForPrep` / `Ready for Prep`; this covers the included workbook where readiness is in column J as well as team workbooks that put a ready marker in column H. Mark ready cells with `X`, `Yes`, `Y`, `True`, or `1`.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Then open:

- Pre-prep: <http://localhost:5000/pre-prep>
- Device Coordinators: <http://localhost:5000/coordinator>

## SharePoint synced workbook

Point the app at a locally synced SharePoint workbook path:

```bash
QUEUE_EXCEL_PATH="/path/to/SharePoint/testing QUEUE.xlsx" python app.py
```

Optional environment variables:

- `QUEUE_STATE_PATH`: JSON file used to store FIFO queue membership. Defaults to `data/queue_state.json`.
- `CLAIM_TIMEOUT_SECONDS`: offer timeout before the app can move to the next coordinator. Defaults to `120`.
