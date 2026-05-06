"""CustomTkinter one-screen MVP for the Daily Queue workflow."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from app_config import APP_NAME, DEFAULT_DB_PATH
from db import QueueDatabase
from excel_sync import import_jobs_from_workbook, write_prep_username, write_qa_done
from models import STATUS_PREP_CLAIMED, STATUS_QA_DONE


class DailyQueueApp(ctk.CTk):
    """Single-window dashboard with import, prep claim, and QA completion."""

    def __init__(self, db_path: str | Path = DEFAULT_DB_PATH):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1180x720")
        self.minsize(1000, 600)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.database = QueueDatabase(db_path)
        self.selected_job_id: int | None = None
        self.job_rows: dict[int, ctk.CTkFrame] = {}

        self.initials_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Import a Daily Queue workbook to begin.")
        self.selected_var = tk.StringVar(value="No job selected")

        self._build_layout()
        self.refresh_jobs()

    def _build_layout(self) -> None:
        """Create the pinned toolbar, selected-job bar, and scrollable dashboard."""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        toolbar = ctk.CTkFrame(self, corner_radius=0)
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.grid_columnconfigure(4, weight=1)

        ctk.CTkLabel(toolbar, text="User initials").grid(row=0, column=0, padx=(16, 6), pady=12)
        ctk.CTkEntry(toolbar, textvariable=self.initials_var, width=90, placeholder_text="JP").grid(
            row=0, column=1, padx=6, pady=12
        )
        ctk.CTkButton(toolbar, text="Import from Excel", command=self.import_from_excel).grid(
            row=0, column=2, padx=6, pady=12
        )
        ctk.CTkButton(toolbar, text="Refresh", command=self.refresh_jobs, fg_color="#596579").grid(
            row=0, column=3, padx=6, pady=12
        )
        ctk.CTkLabel(toolbar, textvariable=self.status_var, anchor="e").grid(
            row=0, column=4, sticky="ew", padx=16, pady=12
        )

        action_bar = ctk.CTkFrame(self, corner_radius=0, fg_color="#e9eef6")
        action_bar.grid(row=1, column=0, sticky="ew")
        action_bar.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(action_bar, textvariable=self.selected_var, anchor="w", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="ew", padx=16, pady=10
        )
        ctk.CTkButton(action_bar, text="Claim for Prep", command=self.claim_selected_for_prep).grid(
            row=0, column=1, padx=6, pady=10
        )
        ctk.CTkButton(action_bar, text="Mark QA Done", command=self.mark_selected_qa_done, fg_color="#2f855a").grid(
            row=0, column=2, padx=(6, 16), pady=10
        )

        self.dashboard = ctk.CTkScrollableFrame(self, label_text="Queue Dashboard")
        self.dashboard.grid(row=2, column=0, sticky="nsew", padx=12, pady=12)
        self.dashboard.grid_columnconfigure(0, weight=1)

    def _require_initials(self) -> str | None:
        """Validate initials before a workflow action writes DB or Excel state."""
        initials = self.initials_var.get().strip().upper()
        if not initials:
            messagebox.showwarning("Initials required", "Enter your initials before claiming or completing work.")
            return None
        if len(initials) > 8:
            messagebox.showwarning("Initials too long", "Use short initials, such as JP or CH.")
            return None
        self.initials_var.set(initials)
        return initials

    def import_from_excel(self) -> None:
        """Prompt for a workbook, import queue rows, and refresh the dashboard."""
        workbook = filedialog.askopenfilename(
            title="Choose Daily Queue workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not workbook:
            return
        try:
            jobs = import_jobs_from_workbook(workbook)
            imported = self.database.upsert_imported_jobs(jobs)
        except Exception as exc:  # UI boundary: show file/Excel errors to the user.
            messagebox.showerror("Import failed", str(exc))
            return
        self.status_var.set(f"Imported {imported} queue row(s) from {Path(workbook).name}.")
        self.refresh_jobs()

    def refresh_jobs(self) -> None:
        """Reload all jobs from SQLite and redraw the grouped dashboard."""
        for child in self.dashboard.winfo_children():
            child.destroy()
        self.job_rows.clear()

        jobs = self.database.list_jobs()
        if not jobs:
            ctk.CTkLabel(
                self.dashboard,
                text="No queue jobs yet. Use Import from Excel to load the Daily Queue workbook.",
                text_color="#596579",
            ).grid(row=0, column=0, padx=20, pady=30)
            return

        current_section = None
        grid_row = 0
        for job in jobs:
            if job["section"] != current_section:
                current_section = job["section"]
                self._add_section_header(current_section, grid_row)
                grid_row += 1
            self._add_job_row(job, grid_row)
            grid_row += 1

    def _add_section_header(self, title: str, row: int) -> None:
        """Add a bold separator matching the workbook's visual sections."""
        frame = ctk.CTkFrame(self.dashboard, fg_color="#26364a", corner_radius=8)
        frame.grid(row=row, column=0, sticky="ew", padx=6, pady=(12, 4))
        ctk.CTkLabel(frame, text=title, text_color="white", font=ctk.CTkFont(size=15, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=8
        )

    def _add_job_row(self, job, row: int) -> None:
        """Add one selectable queue card with visual status indicators."""
        color = self._status_color(job)
        frame = ctk.CTkFrame(self.dashboard, fg_color="white", border_width=2, border_color=color)
        frame.grid(row=row, column=0, sticky="ew", padx=6, pady=4)
        frame.grid_columnconfigure(1, weight=2)
        frame.grid_columnconfigure(2, weight=2)
        frame.grid_columnconfigure(3, weight=1)
        frame.grid_columnconfigure(4, weight=1)
        frame.bind("<Button-1>", lambda _event, job_id=job["id"]: self.select_job(job_id))

        name = f"{job['last_name']}, {job['first_name']}".strip(", ")
        labels = [
            (f"Row {job['excel_row']}", 0),
            (name or "Unnamed client", 1),
            (job["device"] or "No device listed", 2),
            (job["loan_type"] or "", 3),
            (job["queue_date"] or "", 4),
            (self._status_text(job), 5),
        ]
        for text, col in labels:
            label = ctk.CTkLabel(frame, text=text, anchor="w", text_color="#102033")
            label.grid(row=0, column=col, sticky="ew", padx=8, pady=(8, 2))
            label.bind("<Button-1>", lambda _event, job_id=job["id"]: self.select_job(job_id))

        details = ctk.CTkLabel(
            frame,
            text=f"Vocabulary: {job['vocabulary'] or '-'}    Notes: {job['notes'] or '-'}",
            anchor="w",
            text_color="#596579",
        )
        details.grid(row=1, column=1, columnspan=5, sticky="ew", padx=8, pady=(0, 8))
        details.bind("<Button-1>", lambda _event, job_id=job["id"]: self.select_job(job_id))
        self.job_rows[job["id"]] = frame

    def _status_color(self, job) -> str:
        """Return a border color that makes workflow state visible."""
        if job["qa_done_at"]:
            return "#2f855a"
        if job["prep_claimed_by"]:
            return "#d69e2e"
        return "#718096"

    def _status_text(self, job) -> str:
        """Human-readable status text for the dashboard badge column."""
        if job["qa_done_at"]:
            return f"QA Done: {job['qa_done_by']}"
        if job["prep_claimed_by"]:
            return f"Ready for QA / Prep: {job['prep_claimed_by']}"
        return job["status"] or "Imported"

    def select_job(self, job_id: int) -> None:
        """Track the selected job and highlight its row."""
        self.selected_job_id = job_id
        job = self.database.get_job(job_id)
        if job:
            name = f"{job['last_name']}, {job['first_name']}".strip(", ") or "Unnamed client"
            self.selected_var.set(f"Selected row {job['excel_row']}: {name} • {job['device'] or 'No device'}")
        for row_id, frame in self.job_rows.items():
            frame.configure(fg_color="#fff8db" if row_id == job_id else "white")

    def claim_selected_for_prep(self) -> None:
        """Claim the selected job in SQLite, then update Excel column I."""
        initials = self._require_initials()
        if not initials or self.selected_job_id is None:
            messagebox.showinfo("Select a job", "Select a queue row before claiming prep.")
            return
        if not self.database.claim_prep(self.selected_job_id, initials):
            messagebox.showwarning("Already claimed", "This job was already claimed by another user.")
            self.refresh_jobs()
            return
        job = self.database.get_job(self.selected_job_id)
        try:
            write_prep_username(job["workbook_path"], job["sheet_name"], job["excel_row"], initials)
        except Exception as exc:  # UI boundary: the DB claim remains as the source of truth.
            messagebox.showwarning("Excel update failed", f"Prep was claimed in the app, but Excel was not updated: {exc}")
        self.status_var.set(f"Prep claimed for Excel row {job['excel_row']} by {initials}.")
        self.refresh_jobs()

    def mark_selected_qa_done(self) -> None:
        """Mark QA complete in SQLite, then update Excel column J."""
        initials = self._require_initials()
        if not initials or self.selected_job_id is None:
            messagebox.showinfo("Select a job", "Select a queue row before marking QA done.")
            return
        if not self.database.mark_qa_done(self.selected_job_id, initials):
            messagebox.showwarning("Already complete", "QA was already completed for this job.")
            self.refresh_jobs()
            return
        job = self.database.get_job(self.selected_job_id)
        try:
            write_qa_done(job["workbook_path"], job["sheet_name"], job["excel_row"], initials)
        except Exception as exc:  # UI boundary: the DB completion remains as the source of truth.
            messagebox.showwarning("Excel update failed", f"QA was marked done in the app, but Excel was not updated: {exc}")
        self.status_var.set(f"QA completed for Excel row {job['excel_row']} by {initials}.")
        self.refresh_jobs()


if __name__ == "__main__":
    DailyQueueApp().mainloop()
