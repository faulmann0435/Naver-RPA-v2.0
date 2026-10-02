"""Pure product-page logic: grouping, option-form helpers, apply_product_edits, save_product merge."""
import dataclasses

import pandas as pd
import pytest

from core.dictionary import DictionarySettings
from store.base import Author, ConflictError
from store.csv_codec import to_csv_text
from store.dictionary_repo import (
    group_id,
    group_id_of,
    load_dictionary_frame,
    product_rows,
    save_dictionary,
    save_product,
)
from store.memory_store import MemoryStore
from store.rules_repo import DICTIONARY_COLUMNS, DICTIONARY_FILE
from ui.option_logic import (
    FormValues,
    form_issues,
    render_values,
    result_lines,
    row_fields,
    values_from_row,
)
from ui.order_logic import (
    UnmatchedItem,
    item_row,
    register_items,
    suggestion_values,
    unmatched_items,
)
from ui.product_logic import (
    DELETE,
    apply_product_edits,
    change_counts,
    common_vendor,
    dirty_products,
    option_prefix,
    product_issues,
    product_message,
    product_name,
    product_option_rows,
    product_summaries,
    summary_table,
)

USER = Author(name="me@example.com", email="me@example.com")
V = "속초 발주양식"
SETTINGS = DictionarySettings()
G1, G2, G3, G4 = (group_id("1", "사과"), group_id("2", "배"), group_id("3", "가지"), group_id("4", "포도"))


def _frame(rows: list[dict]) -> pd.DataFrame:
    base = {**dict.fromkeys(DICTIONARY_COLUMNS, ""), "channel": "naver", "enabled": "1", "vendor_id": V}
    return pd.DataFrame([{**base, **r} for r in rows], columns=DICTIONARY_COLUMNS)


def _rows() -> list[dict]:
    return [
        {"product_no": "1", "option_key": "a", "product_name_ref": "사과", "option_raw_ref": "A", "display_template": "사과 {수량}개",
         "needs_review": "1", "last_seen_at": "2026-01-01"},
        {"product_no": "1", "option_key": "b", "product_name_ref": "사과", "option_raw_ref": "B",
         "display_template": "사과 {수량*2}개", "needs_review": "1", "last_seen_at": "2026-02-01"},
        {"product_no": "2", "option_key": "", "product_name_ref": "배", "display_template": "배 {수량}개", "needs_review": "0"},
        {"product_no": "3", "option_key": "", "product_name_ref": "가지", "display_template": "가지 {수량}개", "needs_review": "0"},
        {"product_no": "4", "option_key": "", "product_name_ref": "포도", "display_template": "포도 {수량}개", "needs_review": "1"},
    ]


def _store(frame: pd.DataFrame) -> MemoryStore:
    store = MemoryStore()
    store.write_text(DICTIONARY_FILE, to_csv_text(frame), None, "seed")
    return store


# ---------------------------------------------------------------- grouping

def test_product_summaries_group_name_and_sort():
    summaries = product_summaries(_frame(_rows()))
    assert [s.group for s in summaries] == [G1, G4, G3, G2]  # needs review first, then by name
    first = summaries[0]
    assert (first.name, first.product_no, first.option_count, first.review_count) == ("사과", "1", 2, 2)
    assert [s.review_count for s in summaries] == [2, 1, 0, 0]


def test_product_summaries_filters():
    frame = _frame(_rows())
    assert [s.group for s in product_summaries(frame, query="포도")] == [G4]
    assert [s.group for s in product_summaries(frame, query="사과")] == [G1]
    assert {s.group for s in product_summaries(frame, review_only=True)} == {G1, G4}
    assert product_summaries(frame, vendor="없는양식") == []
    assert len(product_summaries(frame, vendor=V)) == 4


def test_summary_table_marks_dirty_products():
    table = summary_table(product_summaries(_frame(_rows())), {G4})
    assert list(table.columns) == ["확인 필요", "옵션 수", "저장 안 됨", "상품명"]
    assert table["저장 안 됨"].tolist() == ["", "✎", "", ""]


def test_name_vendor_and_options_helpers():
    frame = _frame([{**r, "vendor_id": "X" if i == 0 else V} for i, r in enumerate(_rows())])
    rows = product_option_rows(frame, G1)
    assert [r["option_key"] for r in rows] == ["a", "b"] and product_name(rows) == "사과"
    assert common_vendor(rows) == "X" and common_vendor([]) == ""
    assert product_name([]) == ""


def test_option_prefix_is_stable_and_distinct():
    assert option_prefix("1", "a|b") == option_prefix("1", "a|b")
    assert option_prefix("1", "a") != option_prefix("1", "a|b") and option_prefix("1", "a") != option_prefix("2", "a")


def test_dirty_products_compares_with_saved_values():
    frame = _frame(_rows())
    row = product_option_rows(frame, G2)[0]
    prefix = option_prefix("2", "")
    same = {prefix: values_from_row(row)}
    assert dirty_products(frame, same, set()) == set()
    changed = {prefix: dataclasses.replace(values_from_row(row), template="배 상자")}
    assert dirty_products(frame, changed, set()) == {G2}
    assert dirty_products(frame, {}, {option_prefix("3", "")}) == {G3}


# ---------------------------------------------------------------- option logic

def test_values_and_row_fields_round_trip():
    values = values_from_row({"display_template": " a {수량} ", "unit_weight_kg": "0.5", "needs_review": "1", "enabled": "1",
                              "vendor_id": V, "sum_group": "문어", "append_to_end": "0"})
    assert values == FormValues("a {수량}", "", V, "문어", 0.5, False, True, True)
    assert row_fields(values)["unit_weight_kg"] == "0.5" and row_fields(values)["needs_review"] == "1"


def test_result_lines_for_templates_sum_groups_and_append():
    plain = FormValues(template="코다리 {수량*8}마리", vendor=V)
    assert result_lines(plain) == [(1, "코다리 8마리"), (2, "코다리 16마리"), (3, "코다리 24마리")]
    assert render_values(FormValues(template="고정", vendor=V), 2) == "고정 (x2)"
    assert render_values(FormValues(template="x {수량}", qty1_template="하나", vendor=V), 1) == "하나"
    group = FormValues(template="무시", sum_group="★피문어", unit_weight_kg=0.5, vendor=V)
    assert [t for _q, t in result_lines(group)] == ["★피문어 500g", "★피문어 1kg", "★피문어 1.5kg"]
    assert render_values(FormValues(template="끝 {수량}", append_to_end=True), 2) == "묶음 끝에 붙음: 끝 2"
    assert render_values(FormValues(template="{수량*x}"), 1) == "(표기 오류)"


def test_form_issues_errors_and_warnings():
    ids = [V]
    assert [i.level for i in form_issues(FormValues(template="고정", vendor=V), ids)] == ["warning"]
    errors = [i for i in form_issues(FormValues(template="", vendor="없음"), ids) if i.level == "error"]
    assert len(errors) == 2
    assert form_issues(FormValues(template="x {수량}", vendor=V), ids) == []
    warn = form_issues(FormValues(template="사과 2개", vendor=V), ids, qty_example=2)
    assert any("등록 당시 수량" in i.message for i in warn)


# ---------------------------------------------------------------- apply_product_edits

def test_apply_product_edits_replaces_and_deletes_only_this_product():
    base = _frame(_rows())
    edit = dataclasses.replace(values_from_row(product_option_rows(base, G1)[0]), template="사과 상자 {수량}", needs_review=False)
    new = apply_product_edits(base, G1, {("1", "a"): edit, ("1", "b"): DELETE}, user="me", now="NOW")
    assert list(new["option_key"]) == ["a", "", "", ""] and len(new) == 4
    first = new.iloc[0]
    assert (first["display_template"], first["needs_review"], first["updated_at"], first["updated_by"]) == ("사과 상자 {수량}", "0", "NOW", "me")
    assert new.iloc[1].equals(base.fillna("").astype(str).iloc[2])  # other products untouched
    assert change_counts(base, new, G1) == (1, 1) and change_counts(base, new, G2) == (0, 0)
    assert product_message("사과", 1, 1) == "품목 수정: 사과 (수정 1, 삭제 1)"


def test_unchanged_values_do_not_stamp_or_modify():
    base = _frame(_rows())
    same = {("1", "a"): values_from_row(product_option_rows(base, G1)[0])}
    new = apply_product_edits(base, G1, same, user="me", now="NOW")
    assert new.equals(base.fillna("").astype(str)) and change_counts(base, new, G1) == (0, 0)


def test_other_products_keys_are_ignored():
    base = _frame(_rows())
    stray = {("2", ""): FormValues(template="바꿈", vendor=V)}
    assert apply_product_edits(base, G1, stray).equals(base.fillna("").astype(str))


def test_product_issues_only_for_this_product():
    rows = _rows()
    rows[2]["vendor_id"] = "없는양식"
    frame = apply_product_edits(
        _frame(rows), G1, {("1", "a"): FormValues(template="", vendor=V), ("1", "b"): FormValues(template="x {수량}", vendor=V)}
    )
    errors, _warnings = product_issues(frame, G1, [V], SETTINGS)
    assert len(errors) == 1 and "사과" in errors[0]
    assert product_issues(frame, G3, [V], SETTINGS)[0] == []


# ---------------------------------------------------------------- save_product

def _edited(base: pd.DataFrame, group: str, key: str, template: str) -> pd.DataFrame:
    row = next(r for r in product_option_rows(base, group) if r["option_key"] == key)
    return apply_product_edits(base, group, {(row["product_no"], key): dataclasses.replace(values_from_row(row), template=template)},
                               user="me", now="NOW")


def test_save_product_without_conflict():
    store = _store(_frame(_rows()))
    base, sha = load_dictionary_frame(store)
    new_sha = save_product(store, base, sha, _edited(base, G2, "", "배 상자 {수량}"), G2, USER, "msg")
    frame, latest = load_dictionary_frame(store)
    assert latest == new_sha and frame.loc[2, "display_template"] == "배 상자 {수량}"
    assert store.history(DICTIONARY_FILE)[0].author == "me@example.com"


def test_save_product_merges_when_another_product_changed():
    store = _store(_frame(_rows()))
    base, sha = load_dictionary_frame(store)
    theirs = _edited(base, G4, "", "포도 다른사람 {수량}")
    save_dictionary(store, theirs, sha, Author("o", "o@example.com"), "other")
    save_product(store, base, sha, _edited(base, G2, "", "배 내수정 {수량}"), G2, USER, "mine")
    frame, _ = load_dictionary_frame(store)
    by_no = dict(zip(frame["product_no"], frame["display_template"], strict=True))
    assert by_no["2"] == "배 내수정 {수량}" and by_no["4"] == "포도 다른사람 {수량}"
    assert len(frame) == 5


def test_save_product_merge_applies_deletes_by_key():
    store = _store(_frame(_rows()))
    base, sha = load_dictionary_frame(store)
    save_dictionary(store, _edited(base, G4, "", "포도 변경 {수량}"), sha, Author("o", "o@example.com"), "other")
    new = apply_product_edits(base, G1, {("1", "b"): DELETE})
    save_product(store, base, sha, new, G1, USER, "mine")
    frame, _ = load_dictionary_frame(store)
    assert list(frame["option_key"]) == ["a", "", "", ""] and "포도 변경 {수량}" in set(frame["display_template"])


def test_save_product_conflicts_when_same_product_changed():
    store = _store(_frame(_rows()))
    base, sha = load_dictionary_frame(store)
    save_dictionary(store, _edited(base, G2, "", "배 남의수정 {수량}"), sha, Author("o", "o@example.com"), "other")
    before = load_dictionary_frame(store)
    with pytest.raises(ConflictError):
        save_product(store, base, sha, _edited(base, G2, "", "배 내수정 {수량}"), G2, USER, "mine")
    assert load_dictionary_frame(store)[1] == before[1]
    assert product_rows(before[0], G2) == product_rows(load_dictionary_frame(store)[0], G2)


# ---------------------------------------------------------------- main product + add-on sharing a product_no

MAIN, ADDON = group_id("10", "코다리"), group_id("10", "명태회무침")


def _addon_rows() -> list[dict]:
    return [
        {"product_no": "10", "option_key": "a", "product_name_ref": "코다리", "display_template": "코다리 {수량}", "needs_review": "1"},
        {"product_no": "10", "option_key": "b", "product_name_ref": "코다리", "display_template": "코다리 큰것 {수량}", "needs_review": "0"},
        {"product_no": "10", "option_key": "x", "product_name_ref": "명태회무침", "display_template": "회무침 {수량}", "needs_review": "1"},
        {"product_no": "11", "option_key": "", "product_name_ref": "", "display_template": "이름없음 {수량}", "needs_review": "0"},
    ]


def test_group_id_helpers():
    assert group_id("10", " 코다리 ") == group_id("10", "코다리") != group_id("10", "명태회무침")
    assert group_id_of({"product_no": "10", "product_name_ref": "코다리"}) == MAIN


def test_addon_with_same_product_no_is_its_own_group():
    frame = _frame(_addon_rows())
    summaries = product_summaries(frame)
    by_group = {s.group: s for s in summaries}
    assert set(by_group) == {MAIN, ADDON, group_id("11", "")}
    assert (by_group[MAIN].name, by_group[MAIN].option_count) == ("코다리", 2)
    assert (by_group[ADDON].name, by_group[ADDON].option_count, by_group[ADDON].product_no) == ("명태회무침", 1, "10")
    assert by_group[group_id("11", "")].name == "11"  # no name -> product_no
    assert [r["display_template"] for r in product_option_rows(frame, ADDON)] == ["회무침 {수량}"]
    assert [s.group for s in product_summaries(frame, query="회무침")] == [ADDON]


def test_dirty_mark_is_per_group_even_with_equal_option_keys():
    frame = _frame(_addon_rows())
    row = product_option_rows(frame, ADDON)[0]
    pending = {option_prefix("10", "x"): dataclasses.replace(values_from_row(row), template="x {수량}")}
    assert dirty_products(frame, pending, set()) == {ADDON}
    table = summary_table(product_summaries(frame), {ADDON})
    assert dict(zip(table["상품명"], table["저장 안 됨"], strict=True))["명태회무침"] == "✎"
    assert dict(zip(table["상품명"], table["저장 안 됨"], strict=True))["코다리"] == ""


def test_apply_and_counts_touch_only_the_addon_group():
    base = _frame(_addon_rows())
    new = _edited(base, ADDON, "x", "회무침 큰것 {수량}")
    assert change_counts(base, new, ADDON) == (1, 0) and change_counts(base, new, MAIN) == (0, 0)
    assert list(new["display_template"])[:2] == ["코다리 {수량}", "코다리 큰것 {수량}"]
    deleted = apply_product_edits(base, ADDON, {("10", "x"): DELETE})
    assert len(deleted) == 3 and change_counts(base, deleted, ADDON) == (0, 1)


def test_save_addon_group_merges_when_main_group_changed():
    store = _store(_frame(_addon_rows()))
    base, sha = load_dictionary_frame(store)
    save_dictionary(store, _edited(base, MAIN, "a", "코다리 남의수정 {수량}"), sha, Author("o", "o@example.com"), "other")
    save_product(store, base, sha, _edited(base, ADDON, "x", "회무침 내수정 {수량}"), ADDON, USER, "mine")
    frame, _ = load_dictionary_frame(store)
    assert list(frame["display_template"]) == ["코다리 남의수정 {수량}", "코다리 큰것 {수량}", "회무침 내수정 {수량}", "이름없음 {수량}"]


def test_save_addon_group_conflicts_when_same_group_changed():
    store = _store(_frame(_addon_rows()))
    base, sha = load_dictionary_frame(store)
    save_dictionary(store, _edited(base, ADDON, "x", "회무침 남의수정 {수량}"), sha, Author("o", "o@example.com"), "other")
    with pytest.raises(ConflictError):
        save_product(store, base, sha, _edited(base, ADDON, "x", "회무침 내수정 {수량}"), ADDON, USER, "mine")


# ---------------------------------------------------------------- order page logic with form values

def _item(name: str = "상품") -> UnmatchedItem:
    return UnmatchedItem("900", "옵션1", name, "옵션1", 2, 3)


def test_register_items_validates_and_appends():
    store = _store(_frame([]))
    item = _item()
    bad = register_items(store, [(item, FormValues(template="", vendor=V))], USER, SETTINGS, [V])
    assert bad.status == "errors" and "[상품 / 옵션1]" in bad.errors[0]
    plain = FormValues(template="고정 문구", vendor=V)
    assert register_items(store, [(item, plain)], USER, SETTINGS, [V]).status == "needs_confirm"
    assert register_items(store, [], USER, SETTINGS, [V]).status == "nothing_selected"
    values = FormValues(template="상품 {수량*2}개", vendor=V, append_to_end=True)
    done = register_items(store, [(item, values)], USER, SETTINGS, [V])
    assert done.status == "saved"
    frame, _ = load_dictionary_frame(store)
    assert frame.loc[0, "display_template"] == "상품 {수량*2}개" and frame.loc[0, "append_to_end"] == "1"
    assert frame.loc[0, "needs_review"] == "1" and frame.loc[0, "updated_by"] == "me@example.com"


def test_item_row_and_suggestion_values():
    row = item_row(_item(), FormValues(template="a {수량}", vendor=V, sum_group="g", unit_weight_kg=0.5))
    assert row["product_no"] == "900" and row["unit_weight_kg"] == "0.5" and row["vendor_id"] == V
    suggestions = pd.DataFrame([{
        "등록": False, "상품명": "n", "옵션정보": "o", "수량 예시": 2, "건수": 3, "발주양식": V, "발주서 표기": "n {수량*8}",
        "수량 1일 때 표기": "", "합산 이름": "", "1개당 무게(kg)": float("nan"), "product_no": "9", "option_key": "k",
    }])
    assert suggestion_values(suggestions, 0) == FormValues(template="n {수량*8}", vendor=V, needs_review=True)
    assert unmatched_items(suggestions)[0] == UnmatchedItem("9", "k", "n", "o", 2, 3)
