"""Local user profile persistence for the Daily Queue desktop app."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from app_config import DEFAULT_PROFILE_PATH


@dataclass(frozen=True)
class UserProfile:
    """The local user identity used for workflow claims on this computer."""

    name: str
    initials: str

    @property
    def display_name(self) -> str:
        """Return a compact name/initials label for the UI header."""
        return f"{self.name} ({self.initials})"


def normalize_initials(initials: str) -> str:
    """Normalize and validate short user initials."""
    cleaned = "".join(initials.strip().upper().split())
    if not cleaned:
        raise ValueError("Initials are required.")
    if len(cleaned) > 8:
        raise ValueError("Initials must be 8 characters or fewer.")
    return cleaned


def normalize_name(name: str) -> str:
    """Normalize and validate a display name."""
    cleaned = " ".join(name.strip().split())
    if not cleaned:
        raise ValueError("Name is required.")
    return cleaned


class ProfileStore:
    """Read and write the profile JSON file stored locally on the workstation."""

    def __init__(self, path: str | Path = DEFAULT_PROFILE_PATH):
        self.path = Path(path)

    def load(self) -> UserProfile | None:
        """Load the saved profile, returning None if it has not been created."""
        if not self.path.exists():
            return None
        data = json.loads(self.path.read_text(encoding="utf-8"))
        return UserProfile(
            name=normalize_name(str(data.get("name", ""))),
            initials=normalize_initials(str(data.get("initials", ""))),
        )

    def save(self, name: str, initials: str) -> UserProfile:
        """Validate and save a profile to the local computer."""
        profile = UserProfile(name=normalize_name(name), initials=normalize_initials(initials))
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
            json.dumps({"name": profile.name, "initials": profile.initials}, indent=2),
            encoding="utf-8",
        )
        return profile
