from db import QueueDatabase


def _job(row=3):
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
    }


def test_import_preserves_claim_state_on_reimport(tmp_path):
    db = QueueDatabase(tmp_path / "queue.sqlite3")
    assert db.upsert_imported_jobs([_job()]) == 1
    job_id = db.list_jobs()[0]["id"]
    assert db.claim_prep(job_id, "JP") is True

    changed = _job()
    changed["device"] = "Grid Pad Go"
    assert db.upsert_imported_jobs([changed]) == 1

    job = db.get_job(job_id)
    assert job["device"] == "Grid Pad Go"
    assert job["prep_claimed_by"] == "JP"


def test_claim_prep_is_single_winner(tmp_path):
    db = QueueDatabase(tmp_path / "queue.sqlite3")
    db.upsert_imported_jobs([_job()])
    job_id = db.list_jobs()[0]["id"]

    assert db.claim_prep(job_id, "JP") is True
    assert db.claim_prep(job_id, "CH") is False
    assert db.get_job(job_id)["prep_claimed_by"] == "JP"


def test_mark_qa_done_only_once(tmp_path):
    db = QueueDatabase(tmp_path / "queue.sqlite3")
    db.upsert_imported_jobs([_job()])
    job_id = db.list_jobs()[0]["id"]

    assert db.mark_qa_done(job_id, "QA") is True
    assert db.mark_qa_done(job_id, "ZZ") is False
    job = db.get_job(job_id)
    assert job["qa_done_by"] == "QA"
    assert job["status"] == "QA Done"
