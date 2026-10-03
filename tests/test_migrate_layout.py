"""Layout migration against MemoryStore (offline; the real store and secrets are never read)."""
import pytest

from store.layout_repo import OUTPUT_LAYOUT_FILE
from store.memory_store import MemoryStore
from tests.regression_harness import CONFIG_PATH
from tools import migrate_layout
from tools.migrate_layout import build_files, verify_round_trip

NOW = "2026-10-03T00:00:00+09:00"


class FakeStore(MemoryStore):
    repo, branch = "owner/data", "dev"


@pytest.fixture()
def store(monkeypatch) -> FakeStore:
    fake = FakeStore()
    monkeypatch.setattr(migrate_layout, "load_store", lambda branch: fake)
    return fake


def test_round_trip_verification_passes():
    verify_round_trip(str(CONFIG_PATH), build_files(str(CONFIG_PATH), "migration", NOW))


def test_round_trip_verification_fails_on_changed_csv():
    files = build_files(str(CONFIG_PATH), "migration", NOW)
    broken = {OUTPUT_LAYOUT_FILE: files[OUTPUT_LAYOUT_FILE].replace("수취인명", "수취인이름", 1)}
    with pytest.raises(AssertionError):
        verify_round_trip(str(CONFIG_PATH), broken)


def test_dry_run_writes_nothing(store, capsys):
    assert migrate_layout.main(["--branch", "dev", "--dry-run"]) == 0
    out = capsys.readouterr().out
    assert "round trip check: OK" in out and "action=create" in out and "nothing written" in out
    assert store.read_text(OUTPUT_LAYOUT_FILE) is None


def test_create_then_skip_then_refuse_then_force(store, capsys):
    assert migrate_layout.main(["--branch", "dev"]) == 0
    created = store.read_text(OUTPUT_LAYOUT_FILE)
    assert created is not None and created.content.startswith("﻿양식명칭,파일명,열")
    assert store.history(OUTPUT_LAYOUT_FILE)[0].message == f"chore(data): migrate {OUTPUT_LAYOUT_FILE} from config.xlsx"

    capsys.readouterr()
    assert migrate_layout.main(["--branch", "dev"]) == 0  # same layout, later timestamp: skip
    assert "action=skip" in capsys.readouterr().out and store.read_text(OUTPUT_LAYOUT_FILE).sha == created.sha

    store.write_text(OUTPUT_LAYOUT_FILE, created.content.replace("수취인명", "받는분", 1), created.sha, "edited in the app")
    edited = store.read_text(OUTPUT_LAYOUT_FILE)
    assert migrate_layout.main(["--branch", "dev"]) == 1  # create-only: never overwrites edited data
    assert "action=refuse" in capsys.readouterr().out and store.read_text(OUTPUT_LAYOUT_FILE).sha == edited.sha

    assert migrate_layout.main(["--branch", "dev", "--force"]) == 0
    assert store.read_text(OUTPUT_LAYOUT_FILE).content.count("수취인명") == created.content.count("수취인명")


def test_align_keeps_stored_text_only_for_same_layout():
    files = build_files(str(CONFIG_PATH), "migration", NOW)
    memory = MemoryStore()
    assert migrate_layout.align_with_existing(memory, files) == files
    memory.write_text(OUTPUT_LAYOUT_FILE, build_files(str(CONFIG_PATH), "old", "2000").get(OUTPUT_LAYOUT_FILE), None, "seed")
    assert migrate_layout.align_with_existing(memory, files)[OUTPUT_LAYOUT_FILE] != files[OUTPUT_LAYOUT_FILE]
