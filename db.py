"""SQLite persistence for the Daily Queue dashboard.

The database layer owns lane state and uses short write transactions for claim
operations so two users cannot claim the same job at the same time, even when
multiple app instances point at the same shared SQLite file.
"""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

from models import (
    SECTION_DEFAULT,
    STATUS_COMPLETED,
    STATUS_NOT_READY,
    STATUS_IN_PREP,
    STATUS_IN_QA,
    STATUS_READY_PREP,
    STATUS_READY_QA,
)


def utc_now_iso() -> str:
    """Return a timezone-aware UTC timestamp for audit columns."""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def status_from_excel(imported_job: dict) -> str:
    """Choose the initial app lane for a newly imported Excel row."""
    ready_flag = str(imported_job.get("ready_for_prep", "")).strip().upper()
    prep_initials = str(imported_job.get("prep_initials", "")).strip()
    qa_done = str(imported_job.get("qa_done", "")).strip()

    if qa_done:
        return STATUS_COMPLETED
    if prep_initials:
        return STATUS_IN_PREP
    if ready_flag == "X":
        return STATUS_READY_PREP
    return STATUS_NOT_READY


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
        """Create or migrate the MVP schema if it does not already exist."""
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
                    ready_for_prep TEXT,
                    prep_initials TEXT,
                    prep_claimed_at TEXT,
                    prep_finished_at TEXT,
                    qa_initials TEXT,
                    qa_claimed_at TEXT,
                    qa_done_by TEXT,
                    qa_done_at TEXT,
                    status TEXT NOT NULL DEFAULT 'Ready Prep',
                    imported_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(workbook_path, sheet_name, excel_row)
                )
                """
            )
            self._ensure_columns(conn)
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS status_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_id INTEGER NOT NULL,
                    from_status TEXT,
                    to_status TEXT NOT NULL,
                    actor_initials TEXT,
                    action TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    FOREIGN KEY(job_id) REFERENCES jobs(id) ON DELETE CASCADE
                )
                """
            )
            conn.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_jobs_dashboard
                ON jobs(status, section, excel_row)
                """
            )
            conn.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_status_history_job
                ON status_history(job_id, created_at)
                """
            )

    def _ensure_columns(self, conn: sqlite3.Connection) -> None:
        """Add missing columns when a previous MVP database is opened."""
        existing = {row["name"] for row in conn.execute("PRAGMA table_info(jobs)")}
        columns = {
            "ready_for_prep": "TEXT",
            "prep_initials": "TEXT",
            "prep_finished_at": "TEXT",
            "qa_initials": "TEXT",
            "qa_claimed_at": "TEXT",
        }
        for name, column_type in columns.items():
            if name not in existing:
                conn.execute(f"ALTER TABLE jobs ADD COLUMN {name} {column_type}")

        # Copy old MVP prep column into the new explicit prep_initials column.
        if "prep_claimed_by" in existing:
            conn.execute(
                """
                UPDATE jobs
                SET prep_initials = COALESCE(NULLIF(prep_initials, ''), prep_claimed_by)
                WHERE prep_claimed_by IS NOT NULL AND prep_claimed_by != ''
                """
            )

        conn.execute(
            """
            UPDATE jobs
            SET status = CASE
                WHEN qa_done_at IS NOT NULL THEN ?
                WHEN status IN ('Prep Claimed', 'Imported') AND prep_initials IS NOT NULL AND prep_initials != '' THEN ?
                WHEN status = 'Imported' AND (ready_for_prep IS NULL OR UPPER(ready_for_prep) = 'X') THEN ?
                ELSE status
            END
            WHERE status IN ('Prep Claimed', 'QA Done', 'Imported')
            """,
            (STATUS_COMPLETED, STATUS_IN_PREP, STATUS_READY_PREP),
        )

    def _add_history(
        self,
        conn: sqlite3.Connection,
        job_id: int,
        from_status: str | None,
        to_status: str,
        actor_initials: str | None,
        action: str,
        now: str,
    ) -> None:
        """Insert one lane movement/audit entry."""
        conn.execute(
            """
            INSERT INTO status_history (
                job_id, from_status, to_status, actor_initials, action, created_at
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (job_id, from_status, to_status, actor_initials, action, now),
        )

    def upsert_imported_jobs(self, jobs: Iterable[dict]) -> int:
        """Insert new Excel rows or refresh descriptive fields on existing rows.

        Existing lane, claim, and QA fields are intentionally preserved so
        re-importing a workbook does not wipe app workflow state.
        """
        now = utc_now_iso()
        count = 0
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            for job in jobs:
                initial_status = status_from_excel(job)
                existing = conn.execute(
                    """
                    SELECT id FROM jobs
                    WHERE workbook_path = ? AND sheet_name = ? AND excel_row = ?
                    """,
                    (job["workbook_path"], job["sheet_name"], job["excel_row"]),
                ).fetchone()
                cursor = conn.execute(
                    """
                    INSERT INTO jobs (
                        workbook_path, sheet_name, excel_row, section, last_name,
                        first_name, device, loan_type, queue_date, vocabulary,
                        notes, ready_for_prep, prep_initials, qa_done_by,
                        qa_done_at, status, imported_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(workbook_path, sheet_name, excel_row) DO UPDATE SET
                        section=excluded.section,
                        last_name=excluded.last_name,
                        first_name=excluded.first_name,
                        device=excluded.device,
                        loan_type=excluded.loan_type,
                        queue_date=excluded.queue_date,
                        vocabulary=excluded.vocabulary,
                        notes=excluded.notes,
                        ready_for_prep=excluded.ready_for_prep,
                        status=CASE
                            WHEN jobs.status = 'Not Ready' AND excluded.status = 'Ready Prep' THEN excluded.status
                            ELSE jobs.status
                        END,
                        updated_at=excluded.updated_at
                    RETURNING id
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
                        job.get("ready_for_prep"),
                        job.get("prep_initials"),
                        job.get("qa_done"),
                        now if job.get("qa_done") else None,
                        initial_status,
                        now,
                        now,
                    ),
                )
                job_id = cursor.fetchone()["id"]
                if existing is None:
                    self._add_history(conn, job_id, None, initial_status, None, "import", now)
                count += 1
            conn.execute("COMMIT")
        return count

    def list_jobs(self, status: str | None = None) -> list[sqlite3.Row]:
        """Return jobs in workbook order, optionally filtered to one lane status."""
        with self.connect() as conn:
            if status:
                return list(
                    conn.execute(
                        """
                        SELECT * FROM jobs
                        WHERE status = ?
                        ORDER BY CASE WHEN UPPER(section) LIKE '%EXPEDITE%' THEN 0 ELSE 1 END,
                                 workbook_path, sheet_name, excel_row
                        """,
                        (status,),
                    )
                )
            return list(
                conn.execute(
                    """
                    SELECT * FROM jobs
                    ORDER BY CASE WHEN UPPER(section) LIKE '%EXPEDITE%' THEN 0 ELSE 1 END,
                             workbook_path, sheet_name, excel_row
                    """
                )
            )

    def get_job(self, job_id: int) -> sqlite3.Row | None:
        """Fetch one job by database id."""
        with self.connect() as conn:
            return conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()

    def list_history(self, job_id: int) -> list[sqlite3.Row]:
        """Return status history for one job."""
        with self.connect() as conn:
            return list(
                conn.execute(
                    "SELECT * FROM status_history WHERE job_id = ? ORDER BY created_at, id",
                    (job_id,),
                )
            )

    def claim_prep(self, job_id: int, initials: str) -> bool:
        """Atomically claim a Ready Prep job for prep if nobody else has it."""
        now = utc_now_iso()
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT status FROM jobs WHERE id = ?", (job_id,)).fetchone()
            from_status = row["status"] if row else None
            cursor = conn.execute(
                """
                UPDATE jobs
                SET prep_initials = ?, prep_claimed_at = ?, status = ?, updated_at = ?
                WHERE id = ?
                  AND status = ?
                  AND (prep_initials IS NULL OR prep_initials = '')
                """,
                (initials, now, STATUS_IN_PREP, now, job_id, STATUS_READY_PREP),
            )
            if cursor.rowcount == 1:
                self._add_history(conn, job_id, from_status, STATUS_IN_PREP, initials, "claim_prep", now)
            conn.execute("COMMIT")
            return cursor.rowcount == 1

    def finish_prep(self, job_id: int, initials: str) -> bool:
        """Move the user's own in-prep job into the Ready QA lane."""
        now = utc_now_iso()
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            cursor = conn.execute(
                """
                UPDATE jobs
                SET prep_finished_at = ?, status = ?, updated_at = ?
                WHERE id = ? AND status = ? AND prep_initials = ?
                """,
                (now, STATUS_READY_QA, now, job_id, STATUS_IN_PREP, initials),
            )
            if cursor.rowcount == 1:
                self._add_history(conn, job_id, STATUS_IN_PREP, STATUS_READY_QA, initials, "finish_prep", now)
            conn.execute("COMMIT")
            return cursor.rowcount == 1

    def claim_qa(self, job_id: int, initials: str) -> bool:
        """Atomically claim a Ready QA job before finishing QA."""
        now = utc_now_iso()
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            cursor = conn.execute(
                """
                UPDATE jobs
                SET qa_initials = ?, qa_claimed_at = ?, status = ?, updated_at = ?
                WHERE id = ?
                  AND status = ?
                  AND (qa_initials IS NULL OR qa_initials = '')
                """,
                (initials, now, STATUS_IN_QA, now, job_id, STATUS_READY_QA),
            )
            if cursor.rowcount == 1:
                self._add_history(conn, job_id, STATUS_READY_QA, STATUS_IN_QA, initials, "claim_qa", now)
            conn.execute("COMMIT")
            return cursor.rowcount == 1

    def finish_qa(self, job_id: int, initials: str) -> bool:
        """Complete QA for the user's own QA-claimed job."""
        now = utc_now_iso()
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            cursor = conn.execute(
                """
                UPDATE jobs
                SET qa_done_by = ?, qa_done_at = ?, status = ?, updated_at = ?
                WHERE id = ? AND status = ? AND qa_initials = ? AND qa_done_at IS NULL
                """,
                (initials, now, STATUS_COMPLETED, now, job_id, STATUS_IN_QA, initials),
            )
            if cursor.rowcount == 1:
                self._add_history(conn, job_id, STATUS_IN_QA, STATUS_COMPLETED, initials, "finish_qa", now)
            conn.execute("COMMIT")
            return cursor.rowcount == 1
