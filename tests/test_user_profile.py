import json

import pytest

from user_profile import ProfileStore, normalize_initials


def test_profile_store_saves_and_loads_local_json(tmp_path):
    store = ProfileStore(tmp_path / "profile.json")

    saved = store.save("  Jordan   Doe ", " jp ")
    loaded = store.load()

    assert saved.name == "Jordan Doe"
    assert saved.initials == "JP"
    assert loaded == saved
    assert json.loads((tmp_path / "profile.json").read_text())["initials"] == "JP"


def test_initials_are_required():
    with pytest.raises(ValueError):
        normalize_initials("   ")
