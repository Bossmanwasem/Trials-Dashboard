"""CustomTkinter lane dashboard for the Daily Queue workflow."""

from __future__ import annotations

from pathlib import Path
from tkinter import StringVar, filedialog, messagebox

import customtkinter as ctk

from app_config import APP_NAME, DEFAULT_DB_PATH
from db import QueueDatabase
from excel_sync import import_jobs_from_workbook, write_prep_username, write_qa_done
from models import (
    LANE_COMPLETED,
    LANE_IN_PREP,
    LANE_IN_QA,
    LANE_READY_PREP,
    LANE_READY_QA,
    LANE_STATUSES,
    STATUS_COMPLETED,
    STATUS_IN_PREP,
    STATUS_IN_QA,
    STATUS_READY_PREP,
    STATUS_READY_QA,
)
from user_profile import ProfileStore, UserProfile


LANE_ORDER = [LANE_READY_PREP, LANE_IN_PREP, LANE_READY_QA, LANE_IN_QA, LANE_COMPLETED]
STATUS_COLORS = {
    STATUS_READY_PREP: "#718096",
    STATUS_IN_PREP: "#d69e2e",
    STATUS_READY_QA: "#3182ce",
    STATUS_IN_QA: "#805ad5",
    STATUS_COMPLETED: "#2f855a",
}


class ProfileDialog(ctk.CTkToplevel):
    """Modal dialog used on first launch and from Settings/Profile."""

    def __init__(self, parent: "DailyQueueApp", store: ProfileStore, current: UserProfile | None = None):
        super().__init__(parent)
        self.store = store
        self.profile: UserProfile | None = None
        self.title("User Profile")
        self.geometry("460x360")
        self.minsize(460, 360)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.name_entry = ctk.CTkEntry(self, width=300, placeholder_text="Full name")
        self.initials_entry = ctk.CTkEntry(self, width=120, placeholder_text="Initials")
        if current:
            self.name_entry.insert(0, current.name)
            self.initials_entry.insert(0, current.initials)

        ctk.CTkLabel(self, text="Set up your Daily Queue profile", font=ctk.CTkFont(size=18, weight="bold")).pack(
            padx=24, pady=(24, 8)
        )
        ctk.CTkLabel(
            self,
            text="This is saved locally on this computer and is not stored in the shared queue database.",
            wraplength=340,
            text_color="#596579",
        ).pack(padx=24, pady=(0, 16))
        ctk.CTkLabel(self, text="Name").pack(anchor="w", padx=60)
        self.name_entry.pack(padx=24, pady=(2, 10))
        ctk.CTkLabel(self, text="Initials").pack(anchor="w", padx=150)
        self.initials_entry.pack(padx=24, pady=(2, 16))

        button_row = ctk.CTkFrame(self, fg_color="transparent")
        button_row.pack(fill="x", padx=24, pady=(8, 24))
        button_row.grid_columnconfigure(0, weight=1)
        button_row.grid_columnconfigure(1, weight=1)
        if current:
            ctk.CTkButton(button_row, text="Cancel", command=self.destroy, fg_color="#596579").grid(
                row=0, column=0, sticky="ew", padx=(0, 8)
            )
            save_text = "Save Profile"
        else:
            save_text = "Create Profile"
        ctk.CTkButton(button_row, text=save_text, command=self.save_profile).grid(
            row=0,
            column=1 if current else 0,
            columnspan=1 if current else 2,
            sticky="ew",
            padx=(8, 0) if current else 0,
        )

        self.bind("<Return>", lambda _event: self.save_profile())
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.name_entry.focus_set()

    def save_profile(self) -> None:
        """Validate and persist the profile."""
        try:
            self.profile = self.store.save(self.name_entry.get(), self.initials_entry.get())
        except ValueError as exc:
            messagebox.showwarning("Profile required", str(exc))
            return
        self.destroy()

    def on_close(self) -> None:
        """Keep first-launch setup mandatory while allowing edits to be cancelled."""
        if self.profile is None:
            self.destroy()


class DailyQueueApp(ctk.CTk):
    """Single-window dashboard with profile-aware workflow lanes."""

    def __init__(self, db_path: str | Path = DEFAULT_DB_PATH, profile_store: ProfileStore | None = None):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1280x780")
        self.minsize(1100, 650)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.database = QueueDatabase(db_path)
        self.profile_store = profile_store or ProfileStore()
        self.profile: UserProfile | None = self._load_profile()

        self.user_var = StringVar(value="No user profile set")
        self.status_var = StringVar(value="Import a Daily Queue workbook to begin.")
        self.lane_frames: dict[str, ctk.CTkScrollableFrame] = {}

        self._build_layout()
        self._update_user_label()
        self.after(100, self._ensure_profile)
        self.refresh_jobs()

    def _load_profile(self) -> UserProfile | None:
        """Load a local profile, ignoring invalid saved JSON with a setup prompt."""
        try:
            return self.profile_store.load()
        except (OSError, ValueError, KeyError):
            return None

    def _ensure_profile(self) -> None:
        """Prompt on first launch until a local profile is saved."""
        if self.profile is not None:
            return
        while self.profile is None:
            dialog = ProfileDialog(self, self.profile_store)
            self.wait_window(dialog)
            self.profile = dialog.profile
            if self.profile is None:
                if not messagebox.askretrycancel("Profile required", "A name and initials are required to use the queue."):
                    self.destroy()
                    return
        self._update_user_label()

    def _build_layout(self) -> None:
        """Create the pinned toolbar and tabbed workflow lanes."""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        toolbar = ctk.CTkFrame(self, corner_radius=0)
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(toolbar, textvariable=self.user_var, font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=(16, 8), pady=12
        )
        ctk.CTkButton(toolbar, text="Settings/Profile", command=self.edit_profile, width=140).grid(
            row=0, column=1, padx=6, pady=12
        )
        ctk.CTkButton(toolbar, text="Import from Excel", command=self.import_from_excel, width=140).grid(
            row=0, column=2, padx=6, pady=12
        )
        ctk.CTkLabel(toolbar, textvariable=self.status_var, anchor="e").grid(
            row=0, column=3, sticky="ew", padx=16, pady=12
        )
        ctk.CTkButton(toolbar, text="Refresh", command=self.refresh_jobs, fg_color="#596579", width=100).grid(
            row=0, column=4, padx=(6, 16), pady=12
        )

        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=12, pady=12)
        for lane in LANE_ORDER:
            tab = self.tabs.add(lane)
            tab.grid_columnconfigure(0, weight=1)
            frame = ctk.CTkScrollableFrame(tab)
            frame.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
            frame.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)
            self.lane_frames[lane] = frame

    def _update_user_label(self) -> None:
        """Show the active local profile in the pinned header."""
        if self.profile:
            self.user_var.set(f"Active user: {self.profile.display_name}")
        else:
            self.user_var.set("Active user: profile required")

    def edit_profile(self) -> None:
        """Open the local profile settings dialog."""
        dialog = ProfileDialog(self, self.profile_store, self.profile)
        self.wait_window(dialog)
        if dialog.profile:
            self.profile = dialog.profile
            self._update_user_label()
            self.status_var.set(f"Profile updated for {self.profile.display_name}.")
            self.refresh_jobs()

    def _require_profile(self) -> UserProfile | None:
        """Ensure workflow buttons always use saved profile initials."""
        if self.profile is None:
            self._ensure_profile()
        return self.profile

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
        """Reload all jobs from SQLite and redraw every workflow lane."""
        for lane, frame in self.lane_frames.items():
            for child in frame.winfo_children():
                child.destroy()
            jobs = self.database.list_jobs(LANE_STATUSES[lane])
            self._draw_lane(frame, lane, jobs)

    def _draw_lane(self, frame: ctk.CTkScrollableFrame, lane: str, jobs) -> None:
        """Render a lane header and job cards."""
        ctk.CTkLabel(frame, text=f"{lane} ({len(jobs)})", font=ctk.CTkFont(size=18, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 4)
        )
        if not jobs:
            ctk.CTkLabel(frame, text="No jobs in this lane.", text_color="#596579").grid(
                row=1, column=0, sticky="w", padx=10, pady=16
            )
            return
        current_section = None
        row_index = 1
        for job in jobs:
            if job["section"] != current_section:
                current_section = job["section"]
                self._add_section_header(frame, current_section, row_index)
                row_index += 1
            self._add_job_card(frame, lane, job, row_index)
            row_index += 1

    def _add_section_header(self, parent: ctk.CTkScrollableFrame, title: str, row: int) -> None:
        """Add a bold separator matching the workbook's visual sections."""
        frame = ctk.CTkFrame(parent, fg_color="#26364a", corner_radius=8)
        frame.grid(row=row, column=0, sticky="ew", padx=6, pady=(12, 4))
        ctk.CTkLabel(frame, text=title, text_color="white", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=7
        )

    def _add_job_card(self, parent: ctk.CTkScrollableFrame, lane: str, job, row: int) -> None:
        """Add one job card with lane-specific action buttons."""
        frame = ctk.CTkFrame(parent, fg_color="white", border_width=2, border_color=STATUS_COLORS[job["status"]])
        frame.grid(row=row, column=0, sticky="ew", padx=6, pady=5)
        frame.grid_columnconfigure(0, weight=1)

        name = f"{job['last_name']}, {job['first_name']}".strip(", ") or "Unnamed client"
        title = f"Row {job['excel_row']} • {name} • {job['device'] or 'No device listed'}"
        ctk.CTkLabel(frame, text=title, anchor="w", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="ew", padx=10, pady=(10, 2)
        )
        details = (
            f"Loan: {job['loan_type'] or '-'}    Queue Date: {job['queue_date'] or '-'}    "
            f"Vocabulary: {job['vocabulary'] or '-'}    Notes: {job['notes'] or '-'}"
        )
        ctk.CTkLabel(frame, text=details, anchor="w", text_color="#596579", wraplength=900).grid(
            row=1, column=0, sticky="ew", padx=10, pady=(0, 4)
        )
        ctk.CTkLabel(frame, text=self._status_text(job), anchor="w", text_color="#102033").grid(
            row=2, column=0, sticky="ew", padx=10, pady=(0, 10)
        )

        action_frame = ctk.CTkFrame(frame, fg_color="transparent")
        action_frame.grid(row=0, column=1, rowspan=3, padx=10, pady=10)
        self._add_lane_actions(action_frame, lane, job)

    def _status_text(self, job) -> str:
        """Human-readable status text for each card."""
        if job["status"] == STATUS_READY_PREP:
            return "Ready to claim: column H is X and column I is blank."
        if job["status"] == STATUS_IN_PREP:
            return f"Prep claimed by {job['prep_initials']} at {job['prep_claimed_at'] or '-'}"
        if job["status"] == STATUS_READY_QA:
            return f"Prep finished by {job['prep_initials']} at {job['prep_finished_at'] or '-'}"
        if job["status"] == STATUS_IN_QA:
            return f"QA claimed by {job['qa_initials']} at {job['qa_claimed_at'] or '-'}"
        return f"QA complete by {job['qa_done_by']} at {job['qa_done_at'] or '-'}"

    def _add_lane_actions(self, parent: ctk.CTkFrame, lane: str, job) -> None:
        """Show only actions allowed for the lane and active profile."""
        profile = self.profile
        initials = profile.initials if profile else None
        if lane == LANE_READY_PREP:
            ctk.CTkButton(parent, text="Claim for Prep", command=lambda job_id=job["id"]: self.claim_prep(job_id)).pack()
        elif lane == LANE_IN_PREP and initials == job["prep_initials"]:
            ctk.CTkButton(parent, text="Finish Prep", command=lambda job_id=job["id"]: self.finish_prep(job_id)).pack()
        elif lane == LANE_IN_PREP:
            ctk.CTkLabel(parent, text=f"Claimed by {job['prep_initials']}", text_color="#596579").pack()
        elif lane == LANE_READY_QA:
            ctk.CTkButton(parent, text="Claim QA", command=lambda job_id=job["id"]: self.claim_qa(job_id)).pack()
        elif lane == LANE_IN_QA and initials == job["qa_initials"]:
            ctk.CTkButton(
                parent,
                text="Finish QA",
                fg_color="#2f855a",
                command=lambda job_id=job["id"]: self.finish_qa(job_id),
            ).pack()
        elif lane == LANE_IN_QA:
            ctk.CTkLabel(parent, text=f"QA claimed by {job['qa_initials']}", text_color="#596579").pack()
        else:
            ctk.CTkLabel(parent, text="Done", text_color="#2f855a").pack()

    def claim_prep(self, job_id: int) -> None:
        """Claim a Ready Prep job with the saved local profile initials."""
        profile = self._require_profile()
        if profile is None:
            return
        if not self.database.claim_prep(job_id, profile.initials):
            messagebox.showwarning("Already claimed", "This job was already claimed by another user.")
            self.refresh_jobs()
            return
        job = self.database.get_job(job_id)
        try:
            write_prep_username(job["workbook_path"], job["sheet_name"], job["excel_row"], profile.initials)
        except Exception as exc:  # UI boundary: the DB claim remains as the source of truth.
            messagebox.showwarning("Excel update failed", f"Prep was claimed in the app, but Excel was not updated: {exc}")
        self.status_var.set(f"Prep claimed for Excel row {job['excel_row']} by {profile.initials}.")
        self.refresh_jobs()

    def finish_prep(self, job_id: int) -> None:
        """Move the active user's prep job into Ready for QA."""
        profile = self._require_profile()
        if profile is None:
            return
        if not self.database.finish_prep(job_id, profile.initials):
            messagebox.showwarning("Cannot finish prep", "Only the user who claimed this prep job can finish it.")
            self.refresh_jobs()
            return
        self.status_var.set("Prep finished. Job moved to Ready for QA.")
        self.refresh_jobs()

    def claim_qa(self, job_id: int) -> None:
        """Claim a Ready QA job with the saved local profile initials."""
        profile = self._require_profile()
        if profile is None:
            return
        if not self.database.claim_qa(job_id, profile.initials):
            messagebox.showwarning("Already claimed", "This QA job was already claimed by another user.")
            self.refresh_jobs()
            return
        self.status_var.set("QA claimed. Job moved to Currently Being QA.")
        self.refresh_jobs()

    def finish_qa(self, job_id: int) -> None:
        """Finish QA and write INITIALS-Yes to the original Excel column J."""
        profile = self._require_profile()
        if profile is None:
            return
        if not self.database.finish_qa(job_id, profile.initials):
            messagebox.showwarning("Cannot finish QA", "Only the user who claimed this QA job can finish it.")
            self.refresh_jobs()
            return
        job = self.database.get_job(job_id)
        try:
            write_qa_done(job["workbook_path"], job["sheet_name"], job["excel_row"], profile.initials)
        except Exception as exc:  # UI boundary: the DB completion remains as the source of truth.
            messagebox.showwarning("Excel update failed", f"QA was finished in the app, but Excel was not updated: {exc}")
        self.status_var.set(f"QA completed for Excel row {job['excel_row']} by {profile.initials}.")
        self.refresh_jobs()


if __name__ == "__main__":
    DailyQueueApp().mainloop()
