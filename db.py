"""SQLite persistence for the Daily Queue dashboard.

The database layer owns job state and uses short write transactions for claim
operations so two users cannot claim the same job at the same time, even when
multiple app instances point at the same shared SQLite file.
"""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

from models import STATUS_IMPORTED, STATUS_PREP_CLAIMED, STATUS_QA_DONE, SECTION_DEFAULT


def utc_now_iso() -> str:
    """Return a timezone-aware UTC timestamp for audit columns."""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class QueueDatabase:
    """Small repository wrapper around the SQLite queue database."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.init_db()

    @contextmanager
    def connect(self):
        """Open a SQLite connection with row dictionaries and safe pragmas."""
        self.path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(self.path, timeout=30, isolation_level=None)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        try:
            yield conn
        finally:
            conn.close()

    def init_db(self) -> None:
        """Create the MVP schema if it does not already exist."""
        with self.connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS jobs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    workbook_path TEXT NOT NULL,
                    sheet_name TEXT NOT NULL,
                    excel_row INTEGER NOT NULL,
                    section TEXT NOT NULL DEFAULT 'Unsectioned',
                    last_name TEXT,
                    first_name TEXT,
                    device TEXT,
                    loan_type TEXT,
                    queue_date TEXT,
                    vocabulary TEXT,
                    notes TEXT,
                    prep_claimed_by TEXT,
                    prep_claimed_at TEXT,
                    qa_done_by TEXT,
                    qa_done_at TEXT,
                    status TEXT NOT NULL DEFAULT 'Imported',
                    imported_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(workbook_path, sheet_name, excel_row)
                )
                """
            )
            conn.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_jobs_dashboard
                ON jobs(status, section, excel_row)
                """
            )

    def upsert_imported_jobs(self, jobs: Iterable[dict]) -> int:
        """Insert new Excel rows or refresh descriptive fields on existing rows.

        Existing claim and QA fields are intentionally preserved so re-importing
        a workbook does not wipe app workflow state.
        """
        now = utc_now_iso()
        count = 0
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            for job in jobs:
                conn.execute(
                    """
                    INSERT INTO jobs (
                        workbook_path, sheet_name, excel_row, section, last_name,
                        first_name, device, loan_type, queue_date, vocabulary,
                        notes, status, imported_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(workbook_path, sheet_name, excel_row) DO UPDATE SET
                        section=excluded.section,
                        last_name=excluded.last_name,
                        first_name=excluded.first_name,
                        device=excluded.device,
                        loan_type=excluded.loan_type,
                        queue_date=excluded.queue_date,
                        vocabulary=excluded.vocabulary,
                        notes=excluded.notes,
                        updated_at=excluded.updated_at
                    """,
                    (
                        job["workbook_path"],
                        job["sheet_name"],
                        job["excel_row"],
                        job.get("section") or SECTION_DEFAULT,
                        job.get("last_name"),
                        job.get("first_name"),
                        job.get("device"),
                        job.get("loan_type"),
                        job.get("queue_date"),
                        job.get("vocabulary"),
                        job.get("notes"),
                        STATUS_IMPORTED,
                        now,
                        now,
                    ),
                )
                count += 1
            conn.execute("COMMIT")
        return count

    def list_jobs(self) -> list[sqlite3.Row]:
        """Return all jobs in workbook order for dashboard rendering."""
        with self.connect() as conn:
            return list(
                conn.execute(
                    """
                    SELECT * FROM jobs
                    ORDER BY workbook_path, sheet_name, excel_row
                    """
                )
            )

    def get_job(self, job_id: int) -> sqlite3.Row | None:
        """Fetch one job by database id."""
        with self.connect() as conn:
            return conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()

    def claim_prep(self, job_id: int, initials: str) -> bool:
        """Atomically claim a job for prep if nobody else has claimed it."""
        now = utc_now_iso()
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            cursor = conn.execute(
                """
                UPDATE jobs
                SET prep_claimed_by = ?, prep_claimed_at = ?, status = ?, updated_at = ?
                WHERE id = ?
                  AND (prep_claimed_by IS NULL OR prep_claimed_by = '')
                  AND qa_done_at IS NULL
                """,
                (initials, now, STATUS_PREP_CLAIMED, now, job_id),
            )
            conn.execute("COMMIT")
            return cursor.rowcount == 1

    def mark_qa_done(self, job_id: int, initials: str) -> bool:
        """Mark QA complete once for a job that has not already been completed."""
        now = utc_now_iso()
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            cursor = conn.execute(
                """
                UPDATE jobs
                SET qa_done_by = ?, qa_done_at = ?, status = ?, updated_at = ?
                WHERE id = ? AND qa_done_at IS NULL
                """,
                (initials, now, STATUS_QA_DONE, now, job_id),
            )
            conn.execute("COMMIT")
            return cursor.rowcount == 1
