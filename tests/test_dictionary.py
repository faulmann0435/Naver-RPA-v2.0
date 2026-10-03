import pandas as pd

from core.dictionary import (
    DEFAULT_WORKERS,
    DictionaryEntry,
    DictionarySettings,
    ItemDictionary,
    render_entry,
)
from store.rules_repo import DICTIONARY_COLUMNS

SETTINGS = DictionarySettings()


def _frame(*rows: dict) -> pd.DataFrame:
    return pd.DataFrame([{**dict.fromkeys(DICTIONARY_COLUMNS, ""), **r} for r in rows], columns=DICTIONARY_COLUMNS)


def _row(**kw: str) -> dict:
    base = {"channel": "naver", "enabled": "1", "product_no": "100", "option_key": "A: b", "vendor_id": "V1",
            "display_template": "T {수량}"}
    return {**base, **kw}


def test_entry_parsing():
    d = ItemDictionary(_frame(_row(
        sum_group="G", unit_weight_kg="0.5", append_to_end="true", needs_review="1",
        display_template_qty1="one", product_name_ref="n", option_raw_ref="o",
    )))
    entry = d.lookup(100, "A: b", SETTINGS)
    assert entry == DictionaryEntry(
        product_no="100", option_key="A: b", vendor_id="V1", display_template="T {수량}",
        display_template_qty1="one", sum_group="G", unit_weight_kg=0.5, append_to_end=True,
        needs_review=True, product_name_ref="n", option_raw_ref="o",
    )
    assert entry.key == ("100", "A: b")


def test_blank_numeric_and_bool_fields():
    entry = ItemDictionary(_frame(_row())).lookup("100", "A: b", SETTINGS)
    assert entry is not None
    assert entry.unit_weight_kg is None and not entry.append_to_end and not entry.needs_review
    bad = ItemDictionary(_frame(_row(unit_weight_kg="abc"))).lookup("100", "A: b", SETTINGS)
    assert bad is not None and bad.unit_weight_kg is None


def test_filters_channel_enabled_vendor():
    d = ItemDictionary(_frame(
        _row(product_no="1"),
        _row(product_no="2", channel=""),
        _row(product_no="3", channel="coupang"),
        _row(product_no="4", enabled="0"),
        _row(product_no="5", enabled="TRUE"),
        _row(product_no="6", vendor_id=""),
    ))
    found = {p for p in "123456" if d.lookup(p, "A: b", SETTINGS)}
    assert found == {"1", "2", "5"}
    assert len(d) == 3


def test_lookup_normalizes_both_sides():
    d = ItemDictionary(_frame(_row(option_key="가리비 선택: ⭐국산  참가리비 1kg")))
    order = "수령일 선택 (도착시간 지정불가): 2월 5일 수령 / 가리비 선택: ⭐️국산 참가리비 1kg"
    assert d.lookup(100.0, order, SETTINGS) is not None
    assert d.lookup("100.0", "가리비 선택: 국산 참가리비 1kg", SETTINGS) is not None
    assert d.lookup(101, order, SETTINGS) is None


def test_duplicates_last_wins():
    d = ItemDictionary(_frame(_row(vendor_id="V1"), _row(vendor_id="V2"), _row(vendor_id="V3")))
    assert d.duplicates == [("100", "A: b")]
    assert len(d) == 1
    found = d.lookup(100, "A: b", SETTINGS)
    assert found is not None and found.vendor_id == "V3"


def test_empty_dictionary():
    d = ItemDictionary.empty()
    assert len(d) == 0 and d.lookup(1, "x", SETTINGS) is None


def test_settings_defaults_and_override():
    assert DictionarySettings.from_json_text("{}") == DictionarySettings()
    s = DictionarySettings.from_json_text('{"item_separator": " | "}')
    assert s.item_separator == " | " and s.ignored_option_groups == DictionarySettings().ignored_option_groups
    assert DictionarySettings.from_json_text('{"ignored_option_groups": ["x"]}').ignored_option_groups == ("x",)


def _entry(template: str, qty1: str = "") -> DictionaryEntry:
    return DictionaryEntry("1", "k", "V", template, qty1)


def test_render_entry():
    assert render_entry(_entry("깐멍게 500gx{수량}개", "깐멍게 500g"), 1) == "깐멍게 500g"
    assert render_entry(_entry("깐멍게 500gx{수량}개", "깐멍게 500g"), 2) == "깐멍게 500gx2개"
    assert render_entry(_entry("깐멍게 500gx{수량}개"), 1) == "깐멍게 500gx1개"
    assert render_entry(_entry("깐멍게 500g"), 2) == "깐멍게 500g (x2)"
    assert render_entry(_entry("깐멍게 500g"), 1) == "깐멍게 500g"


def test_workers_parsing():
    assert DictionarySettings.from_json_text("{}").workers == DEFAULT_WORKERS
    assert DictionarySettings().workers == ("사장님", "사장님을노리는님", "개발자")
    parsed = DictionarySettings.from_json_text('{"workers": [" 가 ", "", "나", "가", "  ", "다"]}')
    assert parsed.workers == ("가", "나", "다")
    assert DictionarySettings.from_json_text('{"workers": ["", " "]}').workers == DEFAULT_WORKERS
    assert DictionarySettings.from_json_text('{"workers": []}').workers == DEFAULT_WORKERS
    assert DictionarySettings.from_json_text('{"workers": "x"}').workers == DEFAULT_WORKERS
