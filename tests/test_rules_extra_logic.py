"""Impact comparison, rule test and settings.json helpers of the 고급 설정 page."""
import json

import pandas as pd

from core.dictionary import DictionarySettings, ItemDictionary
from store.rules_repo import load_config_from_store
from tests.test_ui_logic import ORDER, make_store
from ui.rules_extra_logic import (
    build_settings_text,
    clean_groups,
    compare_outputs,
    parse_settings,
    run_impact,
    run_rule_test,
    settings_values,
    validate_settings,
)


def merged(*lines: tuple[str, str]) -> pd.DataFrame:
    return pd.DataFrame({"_VendorID": [v for v, _ in lines], "processed_option": [t for _, t in lines]})


def test_compare_outputs_counts_changed_lines_without_customer_columns():
    before = merged(("A", "문어"), ("B", "홍게 2마리"), ("A", "오징어"))
    after = merged(("A", "문어"), ("B", "홍게 4마리"), ("Unclassified", "오징어"))
    result = compare_outputs(before, after)
    assert (result.total_lines, result.changed_lines) == (3, 2)
    assert result.table.columns.tolist() == ["이전 양식", "이전 품목", "이후 양식", "이후 품목"]
    assert result.table.iloc[0].tolist() == ["B", "홍게 2마리", "B", "홍게 4마리"]
    assert compare_outputs(before, before).changed_lines == 0 and compare_outputs(None, None).total_lines == 0


def test_run_impact_same_rules_changes_nothing_and_edited_rules_change_lines():
    store = make_store()
    saved = load_config_from_store(store, "config.xlsx", "1111")
    same = run_impact(ORDER, saved, saved, ItemDictionary.empty(), DictionarySettings())
    assert same.total_lines > 0 and same.changed_lines == 0
    edited = {**saved, "ProductRoute": saved["ProductRoute"].assign(TargetVendorID="문어 발주 양식")}
    changed = run_impact(ORDER, saved, edited, ItemDictionary.empty(), DictionarySettings())
    assert changed.changed_lines >= 1


def test_rule_test_shows_vendor_text_and_korean_log():
    config = load_config_from_store(make_store(), "config.xlsx", "1111")
    result = run_rule_test("문어 세트", "상품 선택: 피문어 2마리", 2, config)
    assert result is not None and result.vendor and result.text
    assert result.steps and any(s.applied for s in result.steps) and any(not s.applied for s in result.steps)
    assert all(s.note for s in result.steps) and all(s.number >= 1 for s in result.steps)
    assert run_rule_test("", "", 1, config) is None


def test_settings_helpers_keep_layout_and_unknown_keys():
    text = json.dumps({"ignored_option_groups": ["a"], "item_separator": " / ", "x": 1}, ensure_ascii=False, indent=2) + "\n"
    data = parse_settings(text)
    assert settings_values(data) == (" / ", ["a"])
    assert build_settings_text(data, " / ", ["a"]) == text  # unchanged -> identical
    new = json.loads(build_settings_text(data, ", ", ["a", "b"]))
    assert new == {"ignored_option_groups": ["a", "b"], "item_separator": ", ", "x": 1}
    assert parse_settings(None) == {} and parse_settings("[1]") == {}
    assert settings_values({}) == (DictionarySettings().item_separator, list(DictionarySettings().ignored_option_groups))
    assert clean_groups([" a ", "", None, "a", float("nan"), "b"]) == ["a", "b"]
    assert validate_settings("  ") and not validate_settings(" + ")
