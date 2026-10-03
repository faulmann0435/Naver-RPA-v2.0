"""Layout validators: every error and warning, references and the final guard (offline)."""
from collections import Counter

import pandas as pd

from store.layout_repo import LAYOUT_HEADERS, LAYOUT_META, layout_sheet_to_csv
from store.layout_validators import (
    LayoutReferences,
    collect_references,
    layout_final_guard,
    layout_form_names,
    validate_layout,
)

SOURCES = {"수취인명", "processed_option"}


def table(*rows: tuple[str, ...]) -> pd.DataFrame:
    """rows: (form, filename, col, header, source, fixed)"""
    frame = pd.DataFrame(list(rows), columns=LAYOUT_HEADERS)
    for name in LAYOUT_META:
        frame[name] = "naver" if name == "channel" else ""
    return frame


GOOD = table(
    ("A양식", "A파일", "A", "받는분", "수취인명", ""),
    ("A양식", "A파일", "B", "품목", "processed_option", ""),
    ("B양식", "B파일", "A", "이름", "수취인명", ""),
)


def messages(issues, level: str) -> list[str]:
    return [i.message for i in issues if i.level == level]


def test_clean_layout_has_no_issues():
    assert validate_layout(GOOD, GOOD, LayoutReferences(), SOURCES) == []
    assert layout_form_names(GOOD) == ["A양식", "B양식"]


def test_empty_form_name_and_empty_header():
    bad = table(("", "f", "A", "x", "", ""), ("A양식", "f", "A", "", "", ""))
    issues = validate_layout(bad, None, LayoutReferences())
    assert any("양식명칭이 비어" in m for m in messages(issues, "error"))
    assert any("칸 이름이 비어" in m for m in messages(issues, "error"))


def test_duplicate_header_within_one_form_only():
    bad = table(("A", "f", "A", "이름", "", ""), ("A", "f", "B", "이름", "", ""), ("B", "g", "A", "이름", "", ""))
    errors = [i for i in validate_layout(bad, None, LayoutReferences()) if i.level == "error"]
    assert [i.row for i in errors] == [0, 1] and "중복" in errors[0].message


def test_two_file_names_in_one_form():
    bad = table(("A", "f1", "A", "x", "", ""), ("A", "f2", "B", "y", "", ""))
    assert any("파일명이 칸마다 다릅니다" in m for m in messages(validate_layout(bad, None, LayoutReferences()), "error"))


def test_forbidden_characters_in_file_name():
    bad = table(("A", "a/b:c*?\"<>|\\", "A", "x", "", ""))
    message = next(m for m in messages(validate_layout(bad, None, LayoutReferences()), "error") if "쓸 수 없는" in m)
    assert all(ch in message for ch in "/:*?\"<>|\\")


def test_form_without_columns_is_an_error():
    issues = validate_layout(GOOD, GOOD, LayoutReferences(), empty_forms=["새 양식"])
    assert any("'새 양식'에 칸이 하나도 없습니다" in m for m in messages(issues, "error"))


def test_vanished_form_that_is_referenced_blocks_and_counts():
    refs = LayoutReferences(route=Counter({"B양식": 2}), options=Counter({"B양식": 1}), dictionary=Counter({"B양식": 5}))
    new = GOOD[GOOD["양식명칭"] != "B양식"]
    message = messages(validate_layout(new, GOOD, refs), "error")[0]
    assert "'B양식'" in message and "상품분류(고급 설정) 2줄" in message
    assert "옵션규칙(고급 설정) 1줄" in message and "품목 사전 5개" in message and "이름을 바꾸는 것도" in message


def test_rename_counts_as_vanish_but_unreferenced_vanish_is_fine():
    renamed = GOOD.copy()
    renamed.loc[2, "양식명칭"] = "B양식2"
    refs = LayoutReferences(route=Counter({"B양식": 1}))
    assert validate_layout(renamed, GOOD, refs)[0].level == "error"
    assert validate_layout(renamed, GOOD, LayoutReferences()) == []


def test_collect_references_only_enabled_rows_and_not_all():
    route = "양식명칭,enabled,channel\nA양식,1,naver\nA양식,0,naver\nB양식,1,coupang\n"
    options = "양식명칭,enabled,channel\nALL,1,naver\na양식,1,naver\n"
    dictionary = "vendor_id,enabled,channel\nA양식,1,naver\n,1,naver\n"
    refs = collect_references(route, options, dictionary)
    assert refs.route == Counter({"A양식": 1}) and refs.options == Counter({"a양식": 1})
    assert refs.dictionary == Counter({"A양식": 1})
    options_only = validate_layout(GOOD[GOOD["양식명칭"] != "A양식"], GOOD, refs)
    assert "옵션규칙(고급 설정) 1줄" in options_only[0].message  # name match ignores case like the rules do
    assert collect_references(None, "", None) == LayoutReferences()


def test_warnings_for_both_source_and_fixed_and_unknown_source():
    warn = table(("A", "f", "A", "x", "수취인명", "고정"), ("A", "f", "B", "y", "이상한열", ""))
    issues = validate_layout(warn, None, LayoutReferences(), SOURCES)
    assert [i.level for i in issues] == ["warning", "warning"]
    assert "고정 글자만 쓰이고" in issues[0].message and "'이상한열'" in issues[1].message
    assert validate_layout(warn, None, LayoutReferences(), None)[0].level == "warning"  # unknown not checked


def test_other_channel_rows_are_ignored():
    frame = GOOD.copy()
    frame.loc[0, "channel"] = "coupang"
    assert layout_form_names(frame) == ["A양식", "B양식"]
    assert validate_layout(frame, None, LayoutReferences()) == []


def test_final_guard_ok_and_failure():
    route = "우선순위,키워드,양식명칭\n1,DEFAULT,A양식\n"
    options = "순서,적용대상(상품명),ActionType (명령),양식명칭,Parameter (설정값)\n1,ALL,REMOVE_TEXT,ALL,x\n"
    assert layout_final_guard(route, options, layout_sheet_to_csv(GOOD[LAYOUT_HEADERS], "t", "n")) == []
    empty = layout_sheet_to_csv(GOOD[LAYOUT_HEADERS].iloc[0:0], "t", "n")
    assert "처리할 수 없습니다" in layout_final_guard(route, options, empty)[0].message
