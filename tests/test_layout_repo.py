"""Layout CSV conversion: config.xlsx -> CSV -> raw layout round trip (offline)."""
import pandas as pd
import pandas.testing as pdt
import pytest

from core.config_loader import load_config, normalize_config, read_config_sheets
from store.csv_codec import from_csv_text
from store.layout_repo import (
    LAYOUT_HEADERS,
    LAYOUT_META,
    OUTPUT_LAYOUT_FILE,
    layout_csv_to_raw,
    layout_sheet_to_csv,
    load_layout,
)
from store.memory_store import MemoryStore
from store.rules_repo import load_config_from_store, sheets_to_rule_csvs
from tests.regression_harness import CONFIG_PATH

CONFIG = str(CONFIG_PATH)
NOW = "2026-10-03T00:00:00+09:00"


@pytest.fixture(scope="module")
def sheet() -> pd.DataFrame:
    return read_config_sheets(CONFIG)["OutputLayout"]


@pytest.fixture(scope="module")
def text(sheet: pd.DataFrame) -> str:
    return layout_sheet_to_csv(sheet, "migration", NOW)


def test_csv_keeps_original_headers_and_adds_meta(text, sheet):
    frame = from_csv_text(text, as_text=True)
    assert frame.columns.tolist() == LAYOUT_HEADERS + LAYOUT_META
    assert len(frame) == len(sheet)
    assert set(frame["channel"]) == {"naver"} and set(frame["updated_by"]) == {"migration"}
    assert text.startswith("\ufeff")


def test_raw_layout_equals_sheet(text, sheet):
    pdt.assert_frame_equal(layout_csv_to_raw(text), sheet)


def test_normalized_layout_equals_config_xlsx(text):
    sheets = read_config_sheets(CONFIG)
    from_csv = normalize_config({**sheets, "OutputLayout": layout_csv_to_raw(text)})
    pdt.assert_frame_equal(from_csv["OutputLayout"], load_config(CONFIG)["OutputLayout"])


def test_other_channels_are_filtered_out(text):
    frame = from_csv_text(text, as_text=True)
    frame.loc[0, "channel"] = "coupang"
    from store.csv_codec import to_csv_text

    assert len(layout_csv_to_raw(to_csv_text(frame))) == len(frame) - 1


def test_load_layout_missing_file_is_none():
    assert load_layout(MemoryStore()) is None


def test_load_layout_returns_frame_and_sha(text, sheet):
    store = MemoryStore()
    sha = store.write_text(OUTPUT_LAYOUT_FILE, text, None, "seed")
    snapshot = load_layout(store)
    assert snapshot is not None and snapshot.sha == sha and snapshot.text == text
    pdt.assert_frame_equal(snapshot.raw, sheet)


def test_config_from_store_uses_store_layout_when_present(text):
    files = sheets_to_rule_csvs(read_config_sheets(CONFIG), "t", NOW)
    store = MemoryStore()
    for name, content in files.items():
        store.write_text(name, content, None, "seed")
    before = load_config_from_store(store, CONFIG)
    edited = text.replace("메로 발주양식,메로 발주양식,K,배송메세지", "메로 발주양식,메로 발주양식,K,배송메모")
    assert edited != text
    store.write_text(OUTPUT_LAYOUT_FILE, edited, None, "layout")
    after = load_config_from_store(store, CONFIG)
    assert "배송메모" in set(after["OutputLayout"]["HeaderName"])
    assert "배송메모" not in set(before["OutputLayout"]["HeaderName"])
