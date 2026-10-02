"""History logic: KST times, diff kinds for every file type, revert writes old content + message."""
import json

import pandas as pd
import pytest

from core.config_loader import read_config_sheets
from core.dictionary import DictionarySettings
from store.base import Author, ConflictError
from store.csv_codec import to_csv_text
from store.memory_store import MemoryStore
from store.rules_repo import (
    DICTIONARY_COLUMNS,
    DICTIONARY_FILE,
    OPTION_RULES_FILE,
    PRODUCT_ROUTE_FILE,
    SETTINGS_FILE,
)
from tests.test_ui_logic import make_store
from ui.history_logic import (
    REVIVES,
    VANISHES,
    diff_against_current,
    diff_dictionary,
    diff_rules,
    diff_settings,
    history_table,
    kst_text,
    revert_errors,
    revert_file,
    revert_message,
    revision_label,
)
from ui.rules_logic import OPTIONS_SPEC, ROUTE_SPEC

SETTINGS = DictionarySettings()
AUTHOR = Author("u@example.com", "u@example.com")


def dict_csv(*rows: dict) -> str:
    base = dict.fromkeys(DICTIONARY_COLUMNS, "")
    return to_csv_text(pd.DataFrame([{**base, "channel": "naver", "enabled": "1", **r} for r in rows], columns=DICTIONARY_COLUMNS))


def item(no: str, name: str, **kw: str) -> dict:
    return {"product_no": no, "option_key": "", "product_name_ref": name, "vendor_id": "V", "display_template": name, **kw}


def test_kst_conversion():
    assert kst_text("2026-01-05T00:30:00Z") == "2026-01-05 09:30"
    assert kst_text("2026-01-05T23:30:00+00:00") == "2026-01-06 08:30"
    assert kst_text("not a date") == "not a date"


def test_history_table_and_label():
    store = MemoryStore()
    store.write_text("a", "1", None, "첫 저장", AUTHOR)
    revs = store.history("a")
    assert history_table(revs).columns.tolist() == ["시각", "작업자", "내용"]
    assert revision_label(revs[0]).endswith("u@example.com · 첫 저장") and revs[0].date[:10] == "2000-01-01"


def test_dictionary_diff_kinds():
    old = dict_csv(item("1", "사과"), item("2", "배"), item("3", "감", display_template="감 {수량}"))
    now = dict_csv(item("1", "사과"), item("3", "감", display_template="감 {수량}개", needs_review="1"), item("4", "귤"))
    table = diff_dictionary(old, now, SETTINGS).table
    by = {r["상품명"]: r for _, r in table.iterrows()}
    assert by["귤"]["구분"] == "추가" and VANISHES in by["귤"]["변경 내용"]
    assert by["배"]["구분"] == "삭제" and REVIVES in by["배"]["변경 내용"]
    assert by["감"]["구분"] == "변경"
    assert "발주서 표기: 감 {수량} → 감 {수량}개" in by["감"]["변경 내용"] and "검토 필요: 0 → 1" in by["감"]["변경 내용"]
    assert "사과" not in by and diff_dictionary(old, old, SETTINGS).empty


def test_rules_diff_rows_and_order_only():
    store = make_store()
    route = store.read_text(PRODUCT_ROUTE_FILE).content
    lines = route.splitlines()
    removed = "\n".join([*lines[:3], *lines[4:]]) + "\n"
    table = diff_rules(removed, route, ROUTE_SPEC).table
    assert table["구분"].tolist() == [VANISHES] and "[우선순위 1]" in table.iloc[0]["규칙"]
    back = diff_rules(route, removed, ROUTE_SPEC).table
    assert back["구분"].tolist() == [REVIVES]
    swapped = "\n".join([lines[0], lines[2], lines[1], *lines[3:]]) + "\n"
    result = diff_rules(swapped, route, ROUTE_SPEC)
    assert result.table.empty and "순서" in result.note


def test_option_rules_diff_ignores_renumbering():
    rules = make_store().read_text(OPTION_RULES_FILE).content
    lines = rules.splitlines()
    shifted = [lines[0]]
    for i, line in enumerate(lines[2:], start=1):  # drop the first rule and renumber the rest
        shifted.append(f"{i}," + line.split(",", 1)[1])
    table = diff_rules("\n".join(shifted) + "\n", rules, OPTIONS_SPEC).table
    assert table["구분"].tolist() == [VANISHES]


def test_settings_diff_by_key():
    old = json.dumps({"ignored_option_groups": ["a"], "item_separator": ", ", "extra": 1})
    now = json.dumps({"ignored_option_groups": ["a", "b"], "item_separator": ", "})
    table = diff_settings(old, now).table
    assert table["항목"].tolist() == ["비교 제외 옵션 항목", "extra"]
    assert table.iloc[0]["선택한 버전"] == "a" and table.iloc[0]["현재"] == "a, b"
    assert diff_against_current(SETTINGS_FILE, old, old, SETTINGS).empty


def test_revert_writes_old_content_and_records_message():
    store = MemoryStore()
    first = store.write_text("f.csv", "old", None, "v1", AUTHOR)
    store.write_text("f.csv", "new", first, "v2", AUTHOR)
    current = store.read_text("f.csv")
    rev = next(r for r in store.history("f.csv") if r.message == "v1")
    revert_file(store, "f.csv", rev, current.sha, AUTHOR)
    assert store.read_text("f.csv").content == "old"
    latest = store.history("f.csv")[0]
    assert latest.message == revert_message("f.csv", rev) == f"되돌리기: f.csv → {kst_text(rev.date)} 버전"
    assert len(store.history("f.csv")) == 3  # history keeps everything
    with pytest.raises(ConflictError):
        revert_file(store, "f.csv", rev, current.sha, AUTHOR)  # stale sha


def test_revert_message_uses_korean_file_name():
    store = MemoryStore()
    store.write_text(SETTINGS_FILE, "{}", None, "x", AUTHOR)
    rev = store.history(SETTINGS_FILE)[0]
    assert revert_message(SETTINGS_FILE, rev).startswith("되돌리기: 설정 → ")


def test_revert_errors_dictionary_rules_and_settings():
    layout = read_config_sheets("config.xlsx", "1111")["OutputLayout"]
    store = make_store()
    route = store.read_text(PRODUCT_ROUTE_FILE).content
    rules = store.read_text(OPTION_RULES_FILE).content
    good = dict_csv(item("1", "사과"))
    assert revert_errors(DICTIONARY_FILE, good, ["V"], SETTINGS, "", layout) == []
    assert revert_errors(DICTIONARY_FILE, good, ["다른양식"], SETTINGS, "", layout)  # vendor no longer exists
    vendors = list(dict.fromkeys(str(v).strip() for v in layout.iloc[:, 0].dropna()))
    assert revert_errors(PRODUCT_ROUTE_FILE, route, vendors, SETTINGS, rules, layout) == []
    assert revert_errors(PRODUCT_ROUTE_FILE, "우선순위,x\n", [], SETTINGS, rules, layout)
    assert revert_errors(OPTION_RULES_FILE, rules.splitlines()[0] + "\n", [], SETTINGS, route, layout)
    assert revert_errors(SETTINGS_FILE, "{not json", [], SETTINGS, "", layout)
    assert revert_errors(SETTINGS_FILE, "{}", [], SETTINGS, "", layout) == []


def test_revert_product_route_without_default_is_blocked():
    import pandas as pd

    from core.config_loader import read_config_sheets
    from store.csv_codec import to_csv_text
    from store.rules_repo import OPTION_RULES_FILE as _OR
    from store.rules_repo import PRODUCT_ROUTE_FILE as _PR
    from store.rules_repo import sheets_to_rule_csvs
    from tests.regression_harness import CONFIG_PATH
    from ui.history_logic import revert_errors

    raw = read_config_sheets(str(CONFIG_PATH))
    csvs = sheets_to_rule_csvs(raw, "t", "2026-01-01")
    from store.csv_codec import from_csv_text

    route = from_csv_text(csvs[_PR], as_text=True)
    no_default = route[~route["키워드"].astype(str).str.upper().str.contains("DEFAULT")]
    vendor_ids = list(dict.fromkeys(str(v) for v in raw["OutputLayout"].iloc[:, 0].dropna()))
    from core.dictionary import DictionarySettings

    errors = revert_errors(_PR, to_csv_text(no_default), vendor_ids, DictionarySettings(), csvs[_OR], raw["OutputLayout"])
    assert any("DEFAULT" in e for e in errors)
    assert revert_errors(_PR, csvs[_PR], vendor_ids, DictionarySettings(), csvs[_OR], raw["OutputLayout"]) == []
    assert isinstance(route, pd.DataFrame)
