import json
import os
import threading
from copy import deepcopy
from datetime import datetime, timedelta, timezone
from pathlib import Path
from uuid import uuid4

from flask import Flask, jsonify, redirect, render_template, request, url_for
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL_PATH = BASE_DIR / "testing QUEUE.xlsx"
EXCEL_PATH = Path(os.environ.get("QUEUE_EXCEL_PATH", DEFAULT_EXCEL_PATH)).expanduser()
STATE_PATH = Path(os.environ.get("QUEUE_STATE_PATH", BASE_DIR / "data" / "queue_state.json")).expanduser()
CLAIM_TIMEOUT_SECONDS = int(os.environ.get("CLAIM_TIMEOUT_SECONDS", "120"))

READY_VALUES = {"x", "yes", "y", "true", "1"}
COMPLETE_VALUES = {"x", "yes", "y", "true", "1", "complete", "done"}
PRIORITY_ORDER = ["Expedites", "Funded Rentals", "Accessories", "Requested Ship Dates", "Regular Daily Queue"]

FIELD_ALIASES = {
    "id": ["ID"],
    "last_name": ["LastName", "Last Name"],
    "first_name": ["FirstName", "First Name"],
    "device": ["Device"],
    "loan_type": ["LoanType", "Loan Type"],
    "queue_date": ["QueueDate", "Queue Date"],
    "vocabulary": ["Vocabulary"],
    "notes": ["Notes"],
    "priority": ["Priority"],
    # The seed workbook uses ReadyForPrep in column J. The app also recognizes column H
    # if a team workbook labels H as ReadyForPrep/Ready for Prep.
    "ready_for_prep": ["ReadyForPrep", "Ready for Prep", "Ready", "Prepped?"],
    "prep_claimed_by": ["PrepClaimedBy", "Prep Claimed By"],
    "prep_claimed_at": ["PrepClaimedAt", "Prep Claimed At"],
    "prep_complete": ["PrepComplete", "Prep Complete"],
    "ready_for_qa": ["ReadyForQA", "Ready for QA", "ReadyFor2ndCheck", "Ready for 2nd Check"],
    "qa_claimed_by": ["QAClaimedBy", "QA Claimed By", "SecondCheckClaimedBy"],
    "qa_claimed_at": ["QAClaimedAt", "QA Claimed At", "SecondCheckClaimedAt"],
    "qa_complete": ["QAComplete", "QA Complete", "SecondCheckComplete"],
    "assignment_user": ["AssignmentUser", "Assignment User"],
    "assignment_expires": ["AssignmentExpires", "Assignment Expires"],
    "assignment_status": ["AssignmentStatus", "Assignment Status"],
}

EXCEL_LOCK = threading.Lock()
STATE_LOCK = threading.Lock()


def create_app():
    app = Flask(__name__)

    @app.get("/")
    def index():
        return redirect(url_for("coordinator_dashboard"))

    @app.get("/pre-prep")
    def pre_prep_dashboard():
        return render_template("pre_prep.html", rows=read_rows(), excel_path=str(EXCEL_PATH))

    @app.get("/coordinator")
    def coordinator_dashboard():
        return render_template(
            "coordinator.html",
            rows=build_dashboard_rows(),
            priorities=PRIORITY_ORDER,
            queue_state=load_state(),
            claim_timeout=CLAIM_TIMEOUT_SECONDS,
            excel_path=str(EXCEL_PATH),
        )

    @app.get("/api/dashboard")
    def api_dashboard():
        return jsonify(
            {
                "rows": build_dashboard_rows(),
                "queues": load_state(),
                "claimTimeoutSeconds": CLAIM_TIMEOUT_SECONDS,
            }
        )

    @app.post("/api/preprep/rows")
    def api_add_row():
        payload = request.get_json(force=True)
        row_id = add_excel_row(payload)
        return jsonify({"ok": True, "id": row_id})

    @app.post("/api/queue/<queue_name>/join")
    def api_join_queue(queue_name):
        payload = request.get_json(force=True)
        user = clean_user(payload.get("user"))
        state = join_queue(queue_name, user)
        return jsonify({"ok": True, "queues": state})

    @app.post("/api/queue/<queue_name>/leave")
    def api_leave_queue(queue_name):
        payload = request.get_json(force=True)
        user = clean_user(payload.get("user"))
        state = leave_queue(queue_name, user)
        return jsonify({"ok": True, "queues": state})

    @app.post("/api/order/<int:row_number>/claim")
    def api_claim_order(row_number):
        payload = request.get_json(force=True)
        queue_name = normalize_queue(payload.get("queue", "prep"))
        user = clean_user(payload.get("user"))
        claim = claim_order(row_number, queue_name, user)
        return jsonify(claim), (200 if claim.get("ok") else 409)

    @app.post("/api/order/<int:row_number>/skip")
    def api_skip_order(row_number):
        payload = request.get_json(force=True)
        queue_name = normalize_queue(payload.get("queue", "prep"))
        user = clean_user(payload.get("user"))
        state = rotate_user_to_bottom(queue_name, user)
        update_assignment_status(row_number, "Skipped", user=user)
        return jsonify({"ok": True, "queues": state})

    @app.post("/api/order/<int:row_number>/complete")
    def api_complete_order(row_number):
        payload = request.get_json(force=True)
        queue_name = normalize_queue(payload.get("queue", "prep"))
        user = clean_user(payload.get("user"))
        complete_order(row_number, queue_name, user)
        state = rotate_user_to_bottom(queue_name, user)
        return jsonify({"ok": True, "queues": state})

    @app.post("/api/order/<int:row_number>/assign-next")
    def api_assign_next(row_number):
        payload = request.get_json(force=True)
        queue_name = normalize_queue(payload.get("queue", "prep"))
        assignment = assign_next_user(row_number, queue_name)
        return jsonify(assignment), (200 if assignment.get("ok") else 409)

    return app


def now_iso():
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def clean_user(user):
    user = (user or "").strip()
    if not user:
        raise ValueError("A coordinator name is required.")
    return user


def normalize_queue(queue_name):
    queue_name = (queue_name or "prep").strip().lower()
    if queue_name in {"prep", "qa"}:
        return queue_name
    raise ValueError("Queue must be prep or qa.")


def truthy(value, complete=False):
    if value is None:
        return False
    values = COMPLETE_VALUES if complete else READY_VALUES
    return str(value).strip().lower() in values


def priority_lane(row):
    text = " ".join(
        str(row.get(key) or "")
        for key in ("priority", "loan_type", "notes", "device", "queue_date")
    ).lower()
    if "expedite" in text or "urgent" in text:
        return "Expedites"
    if "funded" in text or "rental" in text:
        return "Funded Rentals"
    if "accessor" in text:
        return "Accessories"
    if "requested ship" in text or "ship date" in text or "rsd" in text:
        return "Requested Ship Dates"
    return "Regular Daily Queue"


def sort_key(row):
    try:
        lane_index = PRIORITY_ORDER.index(row["lane"])
    except ValueError:
        lane_index = len(PRIORITY_ORDER)
    return (lane_index, str(row.get("queue_date") or ""), row.get("row_number", 0))


def build_dashboard_rows():
    rows = read_rows()
    available = []
    for row in rows:
        row["lane"] = priority_lane(row)
        row["prep_ready"] = truthy(row.get("ready_for_prep")) and not truthy(row.get("prep_complete"), complete=True)
        row["qa_ready"] = truthy(row.get("ready_for_qa")) and not truthy(row.get("qa_complete"), complete=True)
        row["assignment_active"] = assignment_active(row)
        if row["prep_ready"] or row["qa_ready"]:
            available.append(row)
    return sorted(available, key=sort_key)


def assignment_active(row):
    expires = row.get("assignment_expires")
    if not expires:
        return False
    try:
        if isinstance(expires, datetime):
            expires_at = expires if expires.tzinfo else expires.replace(tzinfo=timezone.utc)
        else:
            expires_at = datetime.fromisoformat(str(expires).replace("Z", "+00:00"))
        return expires_at > datetime.now(timezone.utc)
    except ValueError:
        return False


def load_workbook_and_map():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Excel queue file not found: {EXCEL_PATH}")
    workbook = load_workbook(EXCEL_PATH)
    worksheet = workbook.active
    headers = {str(cell.value).strip(): cell.column for cell in worksheet[1] if cell.value}
    field_map = {}
    for field, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            if alias in headers:
                field_map[field] = headers[alias]
                break
    return workbook, worksheet, field_map


def read_rows():
    with EXCEL_LOCK:
        workbook, worksheet, field_map = load_workbook_and_map()
        rows = []
        for row_number in range(2, worksheet.max_row + 1):
            if all(worksheet.cell(row_number, col).value in (None, "") for col in range(1, worksheet.max_column + 1)):
                continue
            row = {field: worksheet.cell(row_number, col).value for field, col in field_map.items()}
            column_h_value = worksheet.cell(row_number, 8).value
            if not truthy(row.get("ready_for_prep")) and truthy(column_h_value):
                row["ready_for_prep"] = column_h_value
            row["row_number"] = row_number
            row["display_name"] = f"{row.get('first_name') or ''} {row.get('last_name') or ''}".strip()
            rows.append(row)
        workbook.close()
        return rows


def add_excel_row(payload):
    with EXCEL_LOCK:
        workbook, worksheet, field_map = load_workbook_and_map()
        next_row = worksheet.max_row + 1
        row_id = payload.get("id") or str(uuid4())[:8]
        values = {
            "id": row_id,
            "last_name": payload.get("last_name"),
            "first_name": payload.get("first_name"),
            "device": payload.get("device"),
            "loan_type": payload.get("loan_type"),
            "queue_date": payload.get("queue_date") or datetime.now().date().isoformat(),
            "vocabulary": payload.get("vocabulary"),
            "notes": payload.get("notes"),
            "priority": payload.get("priority"),
            "ready_for_prep": "X" if payload.get("ready_for_prep") else "",
            "prep_complete": "No",
            "ready_for_qa": "No",
            "qa_complete": "No",
            "assignment_status": "Ready" if payload.get("ready_for_prep") else "Draft",
        }
        for field, value in values.items():
            if field in field_map:
                worksheet.cell(next_row, field_map[field]).value = value
        workbook.save(EXCEL_PATH)
        workbook.close()
        return row_id


def set_cell(worksheet, field_map, row_number, field, value):
    if field in field_map:
        worksheet.cell(row_number, field_map[field]).value = value


def update_assignment_status(row_number, status, user=None, queue_name="prep"):
    with EXCEL_LOCK:
        workbook, worksheet, field_map = load_workbook_and_map()
        if user:
            set_cell(worksheet, field_map, row_number, "assignment_user", user)
        set_cell(worksheet, field_map, row_number, "assignment_status", status)
        set_cell(worksheet, field_map, row_number, "assignment_expires", "")
        workbook.save(EXCEL_PATH)
        workbook.close()


def claim_order(row_number, queue_name, user):
    row = next((r for r in build_dashboard_rows() if r["row_number"] == row_number), None)
    if not row:
        return {"ok": False, "message": "Order is no longer available."}
    expected_user = row.get("assignment_user")
    if row.get("assignment_active") and expected_user and expected_user != user:
        return {"ok": False, "message": f"This order is currently offered to {expected_user}."}
    with EXCEL_LOCK:
        workbook, worksheet, field_map = load_workbook_and_map()
        stamp = now_iso()
        if queue_name == "prep":
            set_cell(worksheet, field_map, row_number, "prep_claimed_by", user)
            set_cell(worksheet, field_map, row_number, "prep_claimed_at", stamp)
        else:
            set_cell(worksheet, field_map, row_number, "qa_claimed_by", user)
            set_cell(worksheet, field_map, row_number, "qa_claimed_at", stamp)
        set_cell(worksheet, field_map, row_number, "assignment_user", user)
        set_cell(worksheet, field_map, row_number, "assignment_status", "Claimed")
        set_cell(worksheet, field_map, row_number, "assignment_expires", "")
        workbook.save(EXCEL_PATH)
        workbook.close()
    return {"ok": True, "message": "Order claimed."}


def complete_order(row_number, queue_name, user):
    with EXCEL_LOCK:
        workbook, worksheet, field_map = load_workbook_and_map()
        stamp = now_iso()
        if queue_name == "prep":
            set_cell(worksheet, field_map, row_number, "prep_claimed_by", user)
            set_cell(worksheet, field_map, row_number, "prep_claimed_at", stamp)
            set_cell(worksheet, field_map, row_number, "prep_complete", "Yes")
            set_cell(worksheet, field_map, row_number, "ready_for_qa", "Yes")
            status = "Ready for QA"
        else:
            set_cell(worksheet, field_map, row_number, "qa_claimed_by", user)
            set_cell(worksheet, field_map, row_number, "qa_claimed_at", stamp)
            set_cell(worksheet, field_map, row_number, "qa_complete", "Yes")
            status = "Complete"
        set_cell(worksheet, field_map, row_number, "assignment_user", user)
        set_cell(worksheet, field_map, row_number, "assignment_status", status)
        set_cell(worksheet, field_map, row_number, "assignment_expires", "")
        workbook.save(EXCEL_PATH)
        workbook.close()


def default_state():
    return {"prep": [], "qa": []}


def load_state():
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not STATE_PATH.exists():
        return default_state()
    try:
        with STATE_PATH.open() as fh:
            state = json.load(fh)
    except json.JSONDecodeError:
        state = default_state()
    for queue in ("prep", "qa"):
        state.setdefault(queue, [])
    return state


def save_state(state):
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with STATE_PATH.open("w") as fh:
        json.dump(state, fh, indent=2)
    return deepcopy(state)


def join_queue(queue_name, user):
    queue_name = normalize_queue(queue_name)
    with STATE_LOCK:
        state = load_state()
        state[queue_name] = [entry for entry in state[queue_name] if entry["user"] != user]
        state[queue_name].append({"user": user, "joined_at": now_iso()})
        return save_state(state)


def leave_queue(queue_name, user):
    queue_name = normalize_queue(queue_name)
    with STATE_LOCK:
        state = load_state()
        state[queue_name] = [entry for entry in state[queue_name] if entry["user"] != user]
        return save_state(state)


def rotate_user_to_bottom(queue_name, user):
    queue_name = normalize_queue(queue_name)
    with STATE_LOCK:
        state = load_state()
        remaining = [entry for entry in state[queue_name] if entry["user"] != user]
        remaining.append({"user": user, "joined_at": now_iso()})
        state[queue_name] = remaining
        return save_state(state)


def assign_next_user(row_number, queue_name):
    queue_name = normalize_queue(queue_name)
    state = load_state()
    if not state[queue_name]:
        return {"ok": False, "message": "No coordinators are currently in this queue."}
    row = next((r for r in build_dashboard_rows() if r["row_number"] == row_number), None)
    if not row:
        return {"ok": False, "message": "Order is no longer available."}
    if row.get("assignment_active"):
        return {
            "ok": True,
            "assignedTo": row.get("assignment_user"),
            "expiresAt": row.get("assignment_expires"),
            "message": "Existing offer is still active.",
        }
    previous_user = row.get("assignment_user")
    previous_status = str(row.get("assignment_status") or "")
    if previous_user and previous_status.startswith("Offered to"):
        state = rotate_user_to_bottom(queue_name, previous_user)
        update_assignment_status(row_number, "Expired", user=previous_user)
    if not state[queue_name]:
        return {"ok": False, "message": "No coordinators are currently in this queue."}
    assigned = state[queue_name][0]["user"]
    expires = datetime.now(timezone.utc) + timedelta(seconds=CLAIM_TIMEOUT_SECONDS)
    with EXCEL_LOCK:
        workbook, worksheet, field_map = load_workbook_and_map()
        set_cell(worksheet, field_map, row_number, "assignment_user", assigned)
        set_cell(worksheet, field_map, row_number, "assignment_expires", expires.replace(microsecond=0).isoformat())
        set_cell(worksheet, field_map, row_number, "assignment_status", f"Offered to {assigned}")
        workbook.save(EXCEL_PATH)
        workbook.close()
    return {
        "ok": True,
        "assignedTo": assigned,
        "expiresAt": expires.replace(microsecond=0).isoformat(),
        "message": f"Offer sent to {assigned}.",
    }


app = create_app()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=True)
