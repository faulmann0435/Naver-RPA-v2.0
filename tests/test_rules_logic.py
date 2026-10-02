"""Rules grids: fold / filter / renumber, DEFAULT protection, CSV building and the no-edit regression."""
from io import StringIO

import pandas as pd
import pytest
from pandas.testing import assert_frame_equal

from core.config_loader import read_config_sheets
from store.rules_repo import load_config_from_store, rule_csvs_to_raw_sheets
from tests.test_ui_logic import make_store
from ui.context import vendor_ids
from ui.rules_logic import (
    ALL_FILTER,
    DELETE,
    OPTIONS_SPEC,
    RID,
    ROUTE_SPEC,
    V_ACTION,
    V_DELETE,
    V_KEYWORD,
    V_ORDER,
    V_PARAM,
    V_PRIORITY,
    V_VENDOR,
    build_csv,
    build_view,
    changed_rids,
    filter_option_rids,
    fold_grid,
    judge,
    load_table,
    sort_rids,
    visible_rids,
    with_work,
)

NOW, USER = "2026-02-02T10:00:00+09:00", "u@example.com"
PARAM = "Parameter (설정값)"


def parse(text: str) -> pd.DataFrame:
    return pd.read_csv(StringIO(text.removeprefix("﻿")), dtype=str, keep_default_na=False)


@pytest.fixture
def store():
    return make_store()


@pytest.fixture
def layout():
    return read_config_sheets("config.xlsx", "1111")["OutputLayout"]


@pytest.fixture
def ids(store):
    return vendor_ids(load_config_from_store(store, "config.xlsx", "1111"))


def open_view(state):
    spec = state.spec
    return build_view(state.work, sort_rids(state.work, visible_rids(state.work, spec), spec), spec)


def test_unchanged_save_is_byte_identical(store):
    for spec in (ROUTE_SPEC, OPTIONS_SPEC):
        state = load_table(store, spec)
        assert build_csv(state.base, state.work, spec, USER, NOW) == state.text
        assert changed_rids(state.base, state.work) == ([], [], [])


def test_fold_without_edits_changes_nothing(store):
    for spec in (ROUTE_SPEC, OPTIONS_SPEC):
        state = load_table(store, spec)
        view = open_view(state)
        assert_frame_equal(fold_grid(state.work, view, view.copy(), spec, allow_remove=True), state.work)


def test_regression_unchanged_rules_give_equal_config_frames(store):
    before = load_config_from_store(store, "config.xlsx", "1111")
    for spec in (ROUTE_SPEC, OPTIONS_SPEC):
        state = load_table(store, spec)
        store.write_text(spec.file, build_csv(state.base, state.work, spec, USER, NOW), state.sha, "same")
    after = load_config_from_store(store, "config.xlsx", "1111")
    assert_frame_equal(before["ProductRoute"], after["ProductRoute"])
    assert_frame_equal(before["OptionRules"], after["OptionRules"])


def test_cell_edit_stamps_only_that_row_and_roundtrips(store):
    state = load_table(store, OPTIONS_SPEC)
    view = open_view(state)
    edited = view.copy()
    edited.loc[edited[RID] == 6, V_PARAM] = r"^\d+[\.\)]\s*"
    work = fold_grid(state.work, view, edited, OPTIONS_SPEC, allow_remove=True)
    assert changed_rids(state.base, work) == ([6], [], [])
    text = build_csv(state.base, work, OPTIONS_SPEC, USER, NOW)
    new = parse(text)
    assert new["updated_by"].tolist().count(USER) == 1 and new.loc[6, "updated_at"] == NOW
    sheets = rule_csvs_to_raw_sheets(load_table(store, ROUTE_SPEC).text, text)
    assert sheets["OptionRules"].shape[0] == 60  # enabled rows only, as before


def test_filter_by_vendor_action_and_text(store):
    work = load_table(store, OPTIONS_SPEC).work
    assert len(filter_option_rids(work, ALL_FILTER, ALL_FILTER, "")) == 62
    only = filter_option_rids(work, ALL_FILTER, "FORMAT_QTY", "")
    assert only and all(work.at[r, "ActionType (명령)"] == "FORMAT_QTY" for r in only)
    assert filter_option_rids(work, "ALL", "REMOVE_TEXT", "")
    found = filter_option_rids(work, ALL_FILTER, ALL_FILTER, "특가")
    assert found and all("특가" in work.at[r, PARAM] for r in found)


def test_filtered_edit_and_delete_keep_hidden_rows(store):
    state = load_table(store, OPTIONS_SPEC)
    view = build_view(state.work, filter_option_rids(state.work, ALL_FILTER, "CALC_UNIT", ""), OPTIONS_SPEC)
    edited = view.copy()
    edited.loc[edited.index[0], V_PARAM] = "마리"
    edited.loc[edited.index[1], V_DELETE] = True
    work = fold_grid(state.work, view, edited, OPTIONS_SPEC, allow_remove=False)
    _modified, added, deleted = changed_rids(state.base, work)
    assert (added, deleted) == ([], [int(edited[RID].iloc[1])])
    text = build_csv(state.base, work, OPTIONS_SPEC, USER, NOW)
    assert len(parse(text)) == 61  # exactly one row less, all hidden rows kept


def test_editing_survives_filter_change_via_folded_work(store):
    state = load_table(store, OPTIONS_SPEC)
    first = build_view(state.work, filter_option_rids(state.work, ALL_FILTER, "CALC_UNIT", ""), OPTIONS_SPEC)
    edited = first.copy()
    edited.loc[edited.index[0], V_PARAM] = "수"
    state = with_work(state, fold_grid(state.work, first, edited, OPTIONS_SPEC, False))
    other = build_view(state.work, filter_option_rids(state.work, ALL_FILTER, "REMOVE_TEXT", ""), OPTIONS_SPEC)
    state = with_work(state, fold_grid(state.work, other, other.copy(), OPTIONS_SPEC, False))
    assert state.work.at[int(first[RID].iloc[0]), PARAM] == "수"


def test_added_rows_are_idempotent_and_numbered_last(store):
    state = load_table(store, OPTIONS_SPEC)
    view = open_view(state)
    new = {RID: None, "사용": True, V_ORDER: None, V_VENDOR: "ALL", "적용대상": "ALL", V_ACTION: "APPEND_SUFFIX",
           V_PARAM: "(냉동)", "설명": "", V_DELETE: False}
    edited = pd.concat([view, pd.DataFrame([new])], ignore_index=True)
    once = fold_grid(state.work, view, edited, OPTIONS_SPEC, True)
    twice = fold_grid(state.work, view, edited, OPTIONS_SPEC, True)
    assert_frame_equal(once, twice)
    assert len(once) == 63 and changed_rids(state.base, once) == ([], [62], [])
    last = parse(build_csv(state.base, once, OPTIONS_SPEC, USER, NOW)).iloc[-1]
    assert last[PARAM] == "(냉동)" and last["순서"] == "63" and last["channel"] == "naver" and last["enabled"] == "1"


def test_blank_added_row_is_ignored(store):
    state = load_table(store, ROUTE_SPEC)
    view = open_view(state)
    blank = pd.DataFrame([{RID: None, "사용": True, V_PRIORITY: None, V_KEYWORD: "", V_VENDOR: None}])
    work = fold_grid(state.work, view, pd.concat([view, blank], ignore_index=True), ROUTE_SPEC, True)
    assert len(work) == len(state.work)


def test_renumber_follows_edited_order_with_stable_ties(store):
    state = load_table(store, OPTIONS_SPEC)
    view = open_view(state)
    edited = view.copy()
    edited.loc[edited[RID] == 9, V_ORDER] = 1  # rid 9 ties with rid 0: the original position wins
    work = fold_grid(state.work, view, edited, OPTIONS_SPEC, True)
    new = parse(build_csv(state.base, work, OPTIONS_SPEC, USER, NOW))
    assert new["순서"].tolist() == [str(i) for i in range(1, 63)]
    assert new.loc[0, PARAM] == state.base.loc[0, PARAM] and new.loc[1, PARAM] == state.base.loc[9, PARAM]


def test_route_rows_sorted_by_priority_on_save(store):
    state = load_table(store, ROUTE_SPEC)
    view = open_view(state)
    edited = view.copy()
    edited.loc[edited[RID] == 0, V_PRIORITY] = 5
    work = fold_grid(state.work, view, edited, ROUTE_SPEC, True)
    priorities = parse(build_csv(state.base, work, ROUTE_SPEC, USER, NOW))["우선순위"].astype(int).tolist()
    assert priorities == sorted(priorities) and priorities[-2:] == [5, 999]


def test_default_row_cannot_be_removed(store, ids, layout):
    route, options = load_table(store, ROUTE_SPEC), load_table(store, OPTIONS_SPEC)
    view = open_view(route)
    default_rid = int(route.base.index[route.base["키워드"] == "DEFAULT"][0])
    work = fold_grid(route.work, view, view[view[RID] != default_rid], ROUTE_SPEC, True)
    judged = judge(route, work, ids, options.text, layout, USER, NOW)
    assert any("DEFAULT" in line for line in judged.errors) and judged.deleted == [default_rid]


def test_judge_only_touched_rows_block_and_final_guard_always_applies(store, ids, layout):
    route, options = load_table(store, ROUTE_SPEC), load_table(store, OPTIONS_SPEC)
    untouched = judge(route, route.work, ids, options.text, layout, USER, NOW)
    assert not untouched.has_changes and not untouched.errors
    view = open_view(route)
    edited = view.copy()
    edited.loc[edited[RID] == 7, V_VENDOR] = "없는양식"
    dup = view.copy()
    dup.loc[dup[RID] == 3, V_KEYWORD] = "홍게"  # same keyword as rid 1: a warning on the touched row
    warned = judge(route, fold_grid(route.work, view, dup, ROUTE_SPEC, True), ids, options.text, layout, USER, NOW)
    assert any("중복" in line for line in warned.warnings) and not warned.errors
    bad = judge(route, fold_grid(route.work, view, edited, ROUTE_SPEC, True), ids, options.text, layout, USER, NOW)
    assert any("없는양식" in line for line in bad.errors)
    wiped = options.work.assign(**{DELETE: True})
    guarded = judge(options, wiped, ids, route.text, layout, USER, NOW)
    assert any("처리할 수 없습니다" in line for line in guarded.errors)
