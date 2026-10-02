import pandas as pd

from core.dictionary import DictionarySettings
from store.rules_repo import DICTIONARY_COLUMNS
from store.validators import NO_QTY_WARNING, validate_dictionary, validate_new_entry

VENDORS = ["V1", "V2"]
SETTINGS = DictionarySettings()


def _frame(*rows: dict) -> pd.DataFrame:
    base = {"enabled": "1", "product_no": "1", "option_key": "", "vendor_id": "V1", "display_template": "A {수량}개"}
    return pd.DataFrame(
        [{**dict.fromkeys(DICTIONARY_COLUMNS, ""), **base, **r} for r in rows], columns=DICTIONARY_COLUMNS
    )


def _issues(*rows: dict):
    return validate_dictionary(_frame(*rows), VENDORS, SETTINGS)


def _messages(issues, level="error"):
    return [i.message for i in issues if i.level == level]


def test_valid_row_has_no_issues():
    assert _issues({}) == []


def test_unknown_vendor():
    issues = _issues({"vendor_id": "ZZ"})
    assert issues[0].column == "vendor_id" and issues[0].row == 0


def test_bad_template_syntax_both_columns():
    issues = _issues({"display_template": "{foo}", "display_template_qty1": "{수량"})
    assert {i.column for i in issues if i.level == "error"} == {"display_template", "display_template_qty1"}


def test_empty_template_and_group():
    assert any("모두 비어" in m for m in _messages(_issues({"display_template": ""})))


def test_sum_group_needs_positive_weight():
    assert _messages(_issues({"display_template": "", "sum_group": "G"}))
    assert _messages(_issues({"display_template": "", "sum_group": "G", "unit_weight_kg": "0"}))
    assert _issues({"display_template": "", "sum_group": "G", "unit_weight_kg": "0.5"}) == []


def test_weight_without_group():
    assert any(i.column == "unit_weight_kg" for i in _issues({"unit_weight_kg": "1"}))


def test_sum_group_with_append_to_end():
    issues = _issues({"display_template": "", "sum_group": "G", "unit_weight_kg": "1", "append_to_end": "1"})
    assert [i.column for i in issues] == ["append_to_end"]


def test_duplicates_reported_on_every_row():
    issues = _issues(
        {"option_key": "옵션 ⭐A"}, {"product_no": "2"}, {"product_no": "1.0", "option_key": "옵션 A"}
    )
    dup_rows = sorted(i.row for i in issues if i.column == "option_key")
    assert dup_rows == [0, 2]


def test_ignored_group_makes_duplicate():
    key = "수령일 선택 (도착시간 지정불가): 1일 / 옵션: A"
    issues = _issues({"option_key": key}, {"option_key": "옵션: A"})
    assert len([i for i in issues if i.level == "error"]) == 2


def test_warning_without_qty_placeholder():
    issues = _issues({"display_template": "고정 문구"})
    assert [(i.level, i.message) for i in issues] == [("warning", NO_QTY_WARNING)]


def test_disabled_rows_are_skipped():
    assert _issues({"enabled": "0", "vendor_id": "ZZ"}) == []


def test_new_entry_rules_and_example_quantity_warning():
    row = {"product_no": "1", "vendor_id": "V1", "display_template": "굴 3개"}
    issues = validate_new_entry(row, 3, VENDORS, SETTINGS)
    assert any("등록 당시 수량(3)" in i.message for i in issues if i.level == "warning")
    assert not [i for i in validate_new_entry({**row, "display_template": "굴 {수량}개"}, 3, VENDORS, SETTINGS)]
    assert not any("등록 당시" in i.message for i in validate_new_entry(row, 1, VENDORS, SETTINGS))
    assert not any("등록 당시" in i.message for i in validate_new_entry({**row, "display_template": "굴 13개"}, 3, VENDORS, SETTINGS))
    bad = validate_new_entry({**row, "vendor_id": "ZZ"}, 1, VENDORS, SETTINGS)
    assert bad[0].level == "error" and bad[0].row is None
