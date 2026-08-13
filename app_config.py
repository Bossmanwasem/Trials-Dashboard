"""Application-wide defaults for the Daily Queue desktop app."""

from __future__ import annotations

import os
from pathlib import Path

APP_NAME = "Daily Queue Dashboard"
DEFAULT_DB_PATH = Path("daily_queue.sqlite3")


def default_profile_path() -> Path:
    """Return a per-computer profile path outside the shared SQLite database."""
    base = os.getenv("LOCALAPPDATA")
    if base:
        return Path(base) / "DailyQueueDashboard" / "profile.json"
    return Path.home() / ".daily_queue_dashboard" / "profile.json"


DEFAULT_PROFILE_PATH = default_profile_path()
