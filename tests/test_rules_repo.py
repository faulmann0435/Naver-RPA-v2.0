"""Rule CSV conversion and the store-backed config round trip (offline)."""
import json

import pandas as pd
import pandas.testing as pdt
import pytest

from core.config_loader import load_config, read_config_sheets
from store.base import StoreError
from store.csv_codec import from_csv_text
from store.memory_store import MemoryStore
from store.rules_repo import (
    DICTIONARY_FILE,
    META_COLUMNS,
    OPTION_RULES_FILE,
    PRODUCT_ROUTE_FILE,
    SETTINGS_FILE,
    load_config_from_store,
    load_rules,
    rule_csvs_to_raw_sheets,
    sheets_to_rule_csvs,
)
from tests.regression_harness import CONFIG_PATH

CONFIG = str(CONFIG_PATH)
NOW = "2026-10-02T00:00:00+09:00"


@pytest.fixture(scope="module")
def raw() -> dict[str, pd.DataFrame]:
    return read_config_sheets(CONFIG)


@pytest.fixture(scope="module")
def files(raw: dict[str, pd.DataFrame]) -> dict[str, str]:
    return sheets_to_rule_csvs(raw, migrated_by="migration", now=NOW)


def seeded_store(files: dict[str, str]) -> MemoryStore:
    store = MemoryStore()
    for name, text in files.items():
        store.write_text(name, text, None, "seed")
    return store


def test_row_counts_and_headers(files, raw):
    route = from_csv_text(files[PRODUCT_ROUTE_FILE])
    rules = from_csv_text(files[OPTION_RULES_FILE])
    assert len(route) == 12 and len(rules) == 62
    assert files[PRODUCT_ROUTE_FILE].startswith("﻿")
    src = raw["OptionRules"]
    original = [c for c in src.columns if not (str(c).startswith("Unnamed") and src[c].isna().all())]
    assert rules.columns.tolist() == original + META_COLUMNS
    assert any("명령" in str(c) or "ActionType" in str(c) for c in rules.columns)
    assert len(rules.columns) == len(src.columns) + len(META_COLUMNS)  # stray note column is kept
    assert route.columns.tolist()[-4:] == META_COLUMNS
    assert set(route["channel"]) == {"naver"}


def test_only_unimplemented_actions_disabled(files):
    rules = from_csv_text(files[OPTION_RULES_FILE])
    action_col = next(c for c in rules.columns if "ActionType" in str(c))
    disabled = rules[rules["enabled"] == 0]
    assert sorted(disabled[action_col].str.strip()) == ["MERGE_SUM_WEIGHT", "SET_UNIT_FLAG"]
    assert (from_csv_text(files[PRODUCT_ROUTE_FILE])["enabled"] == 1).all()


def test_dictionary_and_settings(files):
    dictionary = from_csv_text(files[DICTIONARY_FILE])
    assert len(dictionary) == 0 and "product_no" in dictionary.columns
    assert json.loads(files[SETTINGS_FILE]) == {
        "ignored_option_groups": ["수령일 선택 (도착시간 지정불가)"],
        "item_separator": " / ",
    }


def test_round_trip_matches_load_config(files):
    expected = load_config(CONFIG)
    actual = load_config_from_store(seeded_store(files), CONFIG)
    pdt.assert_frame_equal(actual["ProductRoute"], expected["ProductRoute"])
    pdt.assert_frame_equal(actual["OutputLayout"], expected["OutputLayout"])
    assert actual["_debug_OptionRules_renamed_headers"] == expected["_debug_OptionRules_renamed_headers"]
    # Disabled rows are dropped by design; they are no-ops in the engine.
    kept = expected["OptionRules"]
    kept = kept[~kept["ActionType"].str.strip().str.upper().isin({"MERGE_SUM_WEIGHT", "SET_UNIT_FLAG"})]
    pdt.assert_frame_equal(actual["OptionRules"], kept.reset_index(drop=True))


def test_rule_csvs_filter_channel_and_enabled(files):
    route = from_csv_text(files[PRODUCT_ROUTE_FILE])
    route.loc[0, "channel"] = "coupang"
    route.loc[1, "enabled"] = 0
    from store.csv_codec import to_csv_text

    sheets = rule_csvs_to_raw_sheets(to_csv_text(route), files[OPTION_RULES_FILE])
    assert len(sheets["ProductRoute"]) == 10
    assert not set(META_COLUMNS) & set(sheets["ProductRoute"].columns)
    assert sheets["ProductRoute"].index.tolist() == list(range(10))


def test_load_rules_returns_shas(files):
    store = seeded_store(files)
    snap = load_rules(store)
    assert snap.product_route_sha == store.read_text(PRODUCT_ROUTE_FILE).sha
    assert snap.option_rules_sha == store.read_text(OPTION_RULES_FILE).sha
    assert len(snap.option_rules_raw) == 60


def test_missing_rule_file_raises_korean_error():
    with pytest.raises(StoreError, match="규칙 파일"):
        load_rules(MemoryStore())


def test_empty_unnamed_columns_are_dropped():
    sheet = pd.DataFrame({"순서": [1], "ActionType (명령)": ["REMOVE_TEXT"], "Unnamed: 2": [float("nan")]})
    out = sheets_to_rule_csvs({"ProductRoute": sheet, "OptionRules": sheet}, "t", NOW)
    assert "Unnamed" not in out[OPTION_RULES_FILE]
