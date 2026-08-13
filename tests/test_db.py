from db import QueueDatabase
from models import STATUS_COMPLETED, STATUS_IN_PREP, STATUS_IN_QA, STATUS_NOT_READY, STATUS_READY_PREP, STATUS_READY_QA


def _job(row=3, ready_for_prep="X", prep_initials="", qa_done=""):
    return {
        "workbook_path": "C:/Queue/Daily Queue.xlsx",
        "sheet_name": "Daily Queue",
        "excel_row": row,
        "section": "EXPEDITES",
        "last_name": "Doe",
        "first_name": "Jordan",
        "device": "Talk Pad 10",
        "loan_type": "CL",
        "queue_date": "2026-05-06",
        "vocabulary": "Grid",
        "notes": "SC 50 KG",
        "ready_for_prep": ready_for_prep,
        "prep_initials": prep_initials,
        "qa_done": qa_done,
    }


def test_import_sets_ready_prep_from_column_h_x_and_blank_column_i(tmp_path):
    db = QueueDatabase(tmp_path / "queue.sqlite3")
    db.upsert_imported_jobs([_job()])

    job = db.list_jobs()[0]
    assert job["status"] == STATUS_READY_PREP
    assert job["ready_for_prep"] == "X"
    assert job["prep_initials"] == ""


def test_non_x_column_h_import_is_not_ready_until_reimported_ready(tmp_path):
    db = QueueDatabase(tmp_path / "queue.sqlite3")
    db.upsert_imported_jobs([_job(ready_for_prep="")])
    job_id = db.list_jobs()[0]["id"]
    assert db.get_job(job_id)["status"] == STATUS_NOT_READY

    db.upsert_imported_jobs([_job(ready_for_prep="X")])
    assert db.get_job(job_id)["status"] == STATUS_READY_PREP


def test_import_preserves_lane_state_on_reimport(tmp_path):
    db = QueueDatabase(tmp_path / "queue.sqlite3")
    assert db.upsert_imported_jobs([_job()]) == 1
    job_id = db.list_jobs()[0]["id"]
    assert db.claim_prep(job_id, "JP") is True

    changed = _job()
    changed["device"] = "Grid Pad Go"
    assert db.upsert_imported_jobs([changed]) == 1

    job = db.get_job(job_id)
    assert job["device"] == "Grid Pad Go"
    assert job["prep_initials"] == "JP"
    assert job["status"] == STATUS_IN_PREP


def test_claim_prep_is_single_winner_only_from_ready_prep(tmp_path):
    db = QueueDatabase(tmp_path / "queue.sqlite3")
    db.upsert_imported_jobs([_job()])
    job_id = db.list_jobs()[0]["id"]

    assert db.claim_prep(job_id, "JP") is True
    assert db.claim_prep(job_id, "CH") is False
    job = db.get_job(job_id)
    assert job["prep_initials"] == "JP"
    assert job["status"] == STATUS_IN_PREP


def test_finish_prep_and_qa_workflow_requires_claim_owner(tmp_path):
    db = QueueDatabase(tmp_path / "queue.sqlite3")
    db.upsert_imported_jobs([_job()])
    job_id = db.list_jobs()[0]["id"]

    assert db.claim_prep(job_id, "JP") is True
    assert db.finish_prep(job_id, "CH") is False
    assert db.finish_prep(job_id, "JP") is True
    assert db.get_job(job_id)["status"] == STATUS_READY_QA

    assert db.claim_qa(job_id, "QA") is True
    assert db.get_job(job_id)["status"] == STATUS_IN_QA
    assert db.finish_qa(job_id, "ZZ") is False
    assert db.finish_qa(job_id, "QA") is True

    job = db.get_job(job_id)
    assert job["qa_done_by"] == "QA"
    assert job["status"] == STATUS_COMPLETED


def test_status_history_records_lane_movements(tmp_path):
    db = QueueDatabase(tmp_path / "queue.sqlite3")
    db.upsert_imported_jobs([_job()])
    job_id = db.list_jobs()[0]["id"]

    db.claim_prep(job_id, "JP")
    db.finish_prep(job_id, "JP")
    db.claim_qa(job_id, "QA")
    db.finish_qa(job_id, "QA")

    actions = [row["action"] for row in db.list_history(job_id)]
    assert actions == ["import", "claim_prep", "finish_prep", "claim_qa", "finish_qa"]
