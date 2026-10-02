import pandas as pd

from core.dictionary import DictionarySettings
from store.csv_codec import BOM, from_csv_text, to_csv_text
from store.dictionary_repo import load_dictionary, load_settings
from store.memory_store import MemoryStore
from store.rules_repo import DICTIONARY_COLUMNS, DICTIONARY_FILE, SETTINGS_FILE


def _csv(*rows: dict) -> str:
    frame = pd.DataFrame([{**dict.fromkeys(DICTIONARY_COLUMNS, ""), **r} for r in rows], columns=DICTIONARY_COLUMNS)
    return to_csv_text(frame)


def test_missing_files_give_defaults():
    store = MemoryStore()
    dictionary, sha = load_dictionary(store)
    assert len(dictionary) == 0 and sha is None
    assert load_settings(store) == DictionarySettings()


def test_round_trip_and_int_product_no_match():
    store = MemoryStore()
    text = _csv({
        "channel": "naver", "enabled": "1", "product_no": "5568579375", "option_key": "가리비 선택: 국산참가리비",
        "vendor_id": "V1", "display_template": "가리비 {수량}개", "unit_weight_kg": "0.5",
    })
    sha = store.write_text(DICTIONARY_FILE, text, None, "seed")
    dictionary, got_sha = load_dictionary(store)
    assert got_sha == sha and len(dictionary) == 1
    order_option = "수령일 선택 (도착시간 지정불가): 2월 5일 수령 / 가리비 선택: ⭐국산참가리비"
    entry = dictionary.lookup(5568579375, order_option, DictionarySettings())
    assert entry is not None and entry.unit_weight_kg == 0.5 and entry.vendor_id == "V1"


def test_header_only_file_is_empty_dictionary():
    store = MemoryStore()
    store.write_text(DICTIONARY_FILE, to_csv_text(pd.DataFrame(columns=DICTIONARY_COLUMNS)), None, "seed")
    dictionary, sha = load_dictionary(store)
    assert len(dictionary) == 0 and sha is not None


def test_settings_loaded_from_store():
    store = MemoryStore()
    store.write_text(SETTINGS_FILE, '{"item_separator": " | "}', None, "seed")
    assert load_settings(store).item_separator == " | "


def test_from_csv_text_as_text_keeps_strings():
    text = BOM + "a,b\n0012,\n"
    assert from_csv_text(text, as_text=True).iloc[0].tolist() == ["0012", ""]
    assert from_csv_text(text)["a"].iloc[0] == 12


# ---- save / append / conflict retry ------------------------------------------------------
import pytest

from store.base import Author, ConflictError
from store.dictionary_repo import (
    append_entries,
    load_dictionary_frame,
    save_dictionary,
)

AUTHOR = Author(name="tester", email="t@example.com")
ROW = {"product_no": "10", "option_key": "옵션: A", "vendor_id": "V1", "display_template": "A {수량}"}


def _seeded() -> MemoryStore:
    store = MemoryStore()
    store.write_text(DICTIONARY_FILE, _csv({"channel": "naver", "enabled": "1", "product_no": "1",
                                            "vendor_id": "V1", "display_template": "X"}), None, "seed")
    return store


def test_save_dictionary_writes_and_records_author():
    store = _seeded()
    frame, sha = load_dictionary_frame(store)
    frame.loc[0, "display_template"] = "Y"
    new_sha = save_dictionary(store, frame, sha, AUTHOR, "수정")
    assert new_sha != sha
    assert load_dictionary_frame(store)[0].loc[0, "display_template"] == "Y"
    assert store.history(DICTIONARY_FILE)[0].author == "tester"


def test_save_dictionary_stale_sha_conflicts():
    store = _seeded()
    frame, _ = load_dictionary_frame(store)
    with pytest.raises(ConflictError):
        save_dictionary(store, frame, "stale", AUTHOR, "x")


def test_append_entries_defaults_and_skip_existing():
    store = _seeded()
    result = append_entries(store, [ROW, {**ROW, "option_key": "옵션: A"}, {"product_no": "1", "option_key": "",
                                                                           "vendor_id": "V1", "display_template": "Z"}],
                            AUTHOR, DictionarySettings())
    assert result.added == [("10", "옵션: A")] and result.skipped_existing == [("10", "옵션: A"), ("1", "")]
    frame, sha = load_dictionary_frame(store)
    assert sha == result.sha and len(frame) == 2
    row = frame.iloc[1]
    assert (row["channel"], row["enabled"], row["needs_review"], row["source"]) == ("naver", "1", "1", "manual")
    assert row["updated_by"] == "t@example.com" and row["updated_at"].endswith("+09:00")


def test_append_entries_creates_missing_file():
    store = MemoryStore()
    result = append_entries(store, [ROW], AUTHOR, DictionarySettings())
    assert len(result.added) == 1 and len(load_dictionary_frame(store)[0]) == 1


class _RacyStore:
    """First write fails after another writer added an entry (simulated concurrent save)."""

    def __init__(self, inner: MemoryStore, failures: int = 1) -> None:
        self.inner = inner
        self.failures = failures

    def read_text(self, path):
        return self.inner.read_text(path)

    def history(self, path, limit=30):
        return self.inner.history(path, limit)

    def read_text_at(self, path, ref):
        return self.inner.read_text_at(path, ref)

    def write_text(self, path, content, expected_sha, message, author=None):
        if self.failures > 0:
            self.failures -= 1
            append_entries(self.inner, [{**ROW, "product_no": "99"}], AUTHOR, DictionarySettings())
            raise ConflictError("raced")
        return self.inner.write_text(path, content, expected_sha, message, author)


def test_append_entries_retries_and_keeps_both():
    store = _RacyStore(_seeded())
    result = append_entries(store, [ROW], AUTHOR, DictionarySettings())
    keys = set(load_dictionary_frame(store.inner)[0]["product_no"])
    assert keys == {"1", "10", "99"} and result.added == [("10", "옵션: A")]


def test_append_entries_gives_up_after_retries():
    store = _RacyStore(_seeded(), failures=10)
    with pytest.raises(ConflictError):
        append_entries(store, [{**ROW, "product_no": "5"}], AUTHOR, DictionarySettings(), max_retries=2)
