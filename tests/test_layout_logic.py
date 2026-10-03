"""Layout editing logic: view, no-op save, per-form edits, renumbering, META, judging (offline)."""
from collections import Counter

import pandas as pd
import pytest

from core.config_loader import read_config_sheets
from store.csv_codec import from_csv_text
from store.layout_repo import layout_sheet_to_csv
from store.layout_validators import LayoutReferences
from tests.regression_harness import CONFIG_PATH
from ui.layout_logic import (
    BLANK_LABEL,
    KNOWN_SOURCES,
    NEW_COLUMN_LABEL,
    SOURCE_LABELS,
    V_DELETE,
    V_FIXED,
    V_NAME,
    V_ORDER,
    V_SOURCE,
    apply_edit,
    build_view,
    check_new_form,
    column_choices,
    column_letters,
    delete_form,
    fixed_options,
    form_filename,
    form_names,
    judge,
    preview_table,
    read_layout_table,
    set_one_column,
    source_label,
    source_options,
    source_value,
    summary_message,
)

NOW, USER = "2026-10-03T10:00:00+09:00", "사장님"
FORM = "메로 발주양식"
ROUTE = "우선순위,키워드,양식명칭\n1,DEFAULT,메로 발주양식\n"
OPTIONS = "순서,적용대상(상품명),ActionType (명령),양식명칭,Parameter (설정값)\n1,ALL,REMOVE_TEXT,ALL,x\n"


@pytest.fixture(scope="module")
def text() -> str:
    sheet = read_config_sheets(str(CONFIG_PATH))["OutputLayout"]
    return layout_sheet_to_csv(sheet, "migration", "2026-10-02T00:00:00+09:00")


@pytest.fixture()
def table(text) -> pd.DataFrame:
    return read_layout_table(text)


def edit(table, text, grid, filename=None, form=FORM):
    view0 = build_view(table, form)
    name = form_filename(table, form) if filename is None else filename
    return apply_edit(table, text, form, name, view0, grid, USER, NOW)


def lines(text: str) -> list[str]:
    return text.splitlines()


def block_of(result, form=FORM) -> pd.DataFrame:
    frame = from_csv_text(result.text, as_text=True)
    return frame[frame["양식명칭"] == form].reset_index(drop=True)


def new_row(name: str, source: str = BLANK_LABEL, fixed: str = "") -> pd.DataFrame:
    return pd.DataFrame([{V_NAME: name, V_SOURCE: source, V_FIXED: fixed, V_DELETE: False}])


def test_labels_are_unique_and_round_trip():
    assert len(set(SOURCE_LABELS.values())) == len(SOURCE_LABELS)
    assert source_label("processed_option") == "품목 (정리된 옵션)" and source_label("") == BLANK_LABEL
    assert source_label("낯선열") == "낯선열" and source_value("낯선열") == "낯선열"
    assert source_value("받는 분 이름") == "수취인명" and source_value(BLANK_LABEL) == ""
    assert "processed_option" in KNOWN_SOURCES
    assert source_options(["낯선열", "수취인명"])[-1] == "낯선열" and source_options()[0] == BLANK_LABEL


def test_column_letters():
    assert [column_letters(n) for n in (1, 26, 27, 52, 53, 702, 703)] == ["A", "Z", "AA", "AZ", "BA", "ZZ", "AAA"]


def test_view_rows(table):
    view = build_view(table, FORM)
    assert list(view.columns) == ["_rid", V_ORDER, V_NAME, V_SOURCE, V_FIXED, V_DELETE]
    assert view[V_ORDER].tolist() == list(range(1, 12)) and len(view) == 11
    assert view.loc[0, V_NAME] == "수취인명" and view.loc[0, V_SOURCE] == "받는 분 이름"
    assert view.loc[8, V_SOURCE] == "품목 (정리된 옵션)" and view.loc[9, V_FIXED] == "1"
    assert form_names(table)[0] == "프리미엄과메기" and form_filename(table, FORM) == FORM


def test_noop_save_is_identical(table, text):
    result = edit(table, text, build_view(table, FORM))
    assert not result.has_changes and result.text == text and result.changed_positions == ()


def test_blank_added_row_is_ignored(table, text):
    grid = pd.concat([build_view(table, FORM), pd.DataFrame([{V_DELETE: False}])], ignore_index=True)
    assert not edit(table, text, grid).has_changes


def test_rename_one_column_only_touches_that_row_and_meta(table, text):
    grid = build_view(table, FORM)
    grid.loc[1, V_NAME] = "새이름"
    result = edit(table, text, grid)
    assert (result.modified, result.added, result.deleted) == (1, 0, 0) and result.has_changes
    old, new = lines(text), lines(result.text)
    assert len(old) == len(new)
    diff = [i for i, (a, b) in enumerate(zip(old, new)) if a != b]
    assert len(diff) == 1 and "새이름" in new[diff[0]] and USER in new[diff[0]] and NOW in new[diff[0]]
    assert summary_message(FORM, result) == "발주서 양식 수정: 메로 발주양식 (수정 1, 추가 0, 삭제 0)"


def test_other_forms_stay_in_place_and_byte_identical(table, text):
    grid = build_view(table, FORM)
    result = edit(table, text, grid[grid[V_NAME] != "공란3"])
    assert [ln for ln in lines(text) if FORM not in ln] == [ln for ln in lines(result.text) if FORM not in ln]
    first_old = next(i for i, ln in enumerate(lines(text)) if ln.startswith(FORM))
    first_new = next(i for i, ln in enumerate(lines(result.text)) if ln.startswith(FORM))
    assert first_old == first_new


def test_delete_and_renumber_columns(table, text):
    grid = build_view(table, FORM)
    grid.loc[1, V_DELETE] = True  # B, 공란1
    result = edit(table, text, grid)
    block = block_of(result)
    assert block["열"].tolist() == list("ABCDEFGHIJ") and len(block) == 10
    assert result.deleted == 1 and result.modified == 9  # C..K moved one column to the left
    assert block.iloc[0]["updated_by"] == "migration"  # row A did not change
    assert block.iloc[1]["updated_by"] == USER


def test_reorder_by_order_number_ties_keep_grid_order(table, text):
    grid = build_view(table, FORM)
    grid.loc[0, V_ORDER] = 3  # ties with the old 3rd row, which comes earlier in the grid
    block = block_of(edit(table, text, grid))
    assert block["헤더명"].tolist()[:4] == ["공란1", "수취인명", "통합배송지", "수취인연락처1"]
    assert block["열"].tolist() == list("ABCDEFGHIJK")


def test_add_column_change_source_and_fixed(table, text):
    grid = build_view(table, FORM)
    grid.loc[0, V_SOURCE] = "받는 분 연락처"
    grid.loc[1, V_FIXED] = "메모"
    grid = pd.concat([grid, new_row("추가칸", "주문한 분 이름")], ignore_index=True)
    result = edit(table, text, grid)
    block = block_of(result)
    assert (result.modified, result.added, result.deleted) == (2, 1, 0)
    assert block.iloc[0]["매핑데이터"] == "수취인연락처1" and block.iloc[1]["고정값(Hardcoded)"] == "메모"
    last = block.iloc[-1]
    assert (last["열"], last["헤더명"], last["매핑데이터"], last["파일명"], last["channel"]) == ("L", "추가칸", "구매자명", FORM, "naver")
    assert last["updated_by"] == USER and last["updated_at"] == NOW


def test_unknown_source_is_shown_and_preserved(table, text):
    table.loc[table.index[table["양식명칭"] == FORM][0], "매핑데이터"] = "낯선열"
    view = build_view(table, FORM)
    assert view.loc[0, V_SOURCE] == "낯선열"
    changed = view.copy()
    changed.loc[1, V_FIXED] = "x"
    result = apply_edit(table, text, FORM, FORM, view, changed, USER, NOW)
    assert block_of(result).iloc[0]["매핑데이터"] == "낯선열"


def test_file_name_change_marks_every_row(table, text):
    result = edit(table, text, build_view(table, FORM), filename="새파일")
    assert result.modified == 11 and set(block_of(result)["파일명"]) == {"새파일"}


def test_add_form_goes_to_the_end(table, text):
    grid = pd.concat([new_row("받는분", "받는 분 이름"), new_row("주소", "받는 분 주소")], ignore_index=True)
    view0 = build_view(table, "신규")
    assert len(view0) == 0
    result = apply_edit(table, text, "신규", "신규파일", view0, grid, USER, NOW)
    new = from_csv_text(result.text, as_text=True)
    assert new["양식명칭"].tolist()[-2:] == ["신규", "신규"] and new["열"].tolist()[-2:] == ["A", "B"]
    assert lines(result.text)[: len(lines(text))] == lines(text)
    assert result.added == 2 and "신규" in form_names(result.table)


def test_delete_form_removes_all_rows(table, text):
    result = delete_form(table, text, FORM, USER, NOW)
    assert result.deleted == 11 and FORM not in form_names(result.table)
    assert [ln for ln in lines(text) if FORM not in ln] == lines(result.text)


def test_preview_table_final_order_and_example(table):
    grid = build_view(table, FORM)
    grid.loc[0, V_ORDER] = 99
    shown = preview_table(grid)
    assert shown.columns[-1] == "수취인명" and shown.iloc[0, -1] == "(받는 분 이름)"
    assert shown.iloc[0, list(shown.columns).index("수량")] == "1"
    dup = pd.DataFrame({V_NAME: ["가", "가", ""], V_SOURCE: [BLANK_LABEL] * 3, V_FIXED: ["", "", ""], V_DELETE: [False] * 3})
    assert list(preview_table(dup).columns) == ["가", "가 (2)"]


def test_check_new_form():
    assert check_new_form(" ", "f", []) == "양식 이름을 입력하세요."
    assert "이미 있습니다" in check_new_form("A", "f", ["A"])
    assert "쓸 수 없는" in check_new_form("B", "a/b", ["A"])
    assert check_new_form("B", "", ["A"]) is None


def judged(table, text, grid, filename=None, form=FORM, refs=None):
    view0 = build_view(table, form)
    name = form_filename(table, form) if filename is None else filename
    return judge(table, text, form, name, view0, grid, refs or LayoutReferences(), ROUTE, OPTIONS, USER, NOW)


def test_judge_clean_edit_and_noop(table, text):
    assert not judged(table, text, build_view(table, FORM)).has_changes
    grid = build_view(table, FORM)
    grid.loc[1, V_NAME] = "새이름"
    result = judged(table, text, grid)
    assert result.has_changes and result.errors == [] and result.warnings == []


def test_judge_duplicate_header_and_filename_errors(table, text):
    grid = build_view(table, FORM)
    grid.loc[1, V_NAME] = grid.loc[0, V_NAME]
    assert any("중복" in e for e in judged(table, text, grid).errors)
    assert any("쓸 수 없는" in e for e in judged(table, text, build_view(table, FORM), filename="a/b").errors)


def test_judge_warning_only_for_changed_rows(table, text):
    grid = build_view(table, FORM)
    grid.loc[1, V_FIXED] = "고정"
    grid.loc[1, V_SOURCE] = "받는 분 이름"
    grid.loc[2, V_SOURCE] = "낯선열"
    result = judged(table, text, grid)
    assert len(result.warnings) == 2 and result.errors == []


def test_judge_blocks_deleting_a_referenced_form_but_not_an_unused_one(table, text):
    refs = LayoutReferences(route=Counter({FORM: 1}))
    view0 = build_view(table, FORM)
    blocked = judge(table, text, FORM, FORM, view0, view0.iloc[0:0], refs, ROUTE, OPTIONS, USER, NOW)
    assert any("아직 쓰이고 있어" in e for e in blocked.errors)
    other = "홍어 발주양식"
    view1 = build_view(table, other)
    free = judge(table, text, other, other, view1, view1.iloc[0:0], refs, ROUTE, OPTIONS, USER, NOW)
    assert free.errors == []


def test_judge_new_form_without_columns_and_with_columns(table, text):
    view0 = build_view(table, "신규")
    empty = judge(table, text, "신규", "f", view0, view0, LayoutReferences(), ROUTE, OPTIONS, USER, NOW)
    assert not empty.has_changes and any("칸이 하나도 없습니다" in e for e in empty.errors)
    filled = judge(table, text, "신규", "f", view0, new_row("받는분", "받는 분 이름"), LayoutReferences(), ROUTE, OPTIONS, USER, NOW)
    assert filled.has_changes and filled.errors == []


def test_judge_final_guard_when_layout_would_be_empty():
    sheet = pd.DataFrame({"양식명칭": ["A"], "파일명": ["f"], "열": ["A"], "헤더명": ["x"], "매핑데이터": [""], "고정값(Hardcoded)": [""]})
    only = layout_sheet_to_csv(sheet, "m", "n")
    table = read_layout_table(only)
    view0 = build_view(table, "A")
    result = judge(table, only, "A", "f", view0, view0.iloc[0:0], LayoutReferences(), ROUTE, OPTIONS, USER, NOW)
    assert any("처리할 수 없습니다" in e for e in result.errors)


def test_fixed_options_offer_used_texts_most_used_first(table):
    grid = build_view(table, FORM)
    options = fixed_options(table, grid)
    assert options[0] == "" and "최고다농수산" in options and "033-636-0357" in options
    assert len(options) == len(set(options))
    grid.at[0, V_FIXED] = "새 고정 글자"
    assert "새 고정 글자" in fixed_options(table, grid)


def test_column_choices_put_new_first_then_numbered_rows(table):
    grid = build_view(table, FORM)
    choices = column_choices(grid)
    assert choices[0] == NEW_COLUMN_LABEL and len(choices) == len(grid) + 1
    assert choices[1] == f"1. {grid.at[0, V_NAME]}"


def test_set_one_column_changes_a_copy_and_saves_korean(table, text):
    grid = build_view(table, FORM)
    changed = set_one_column(grid, 1, " 받는분 성명 ", SOURCE_LABELS["수취인명"], "")
    assert grid.at[1, V_NAME] != "받는분 성명"  # original untouched
    assert changed.at[1, V_NAME] == "받는분 성명"
    result = edit(table, text, changed)
    assert result.modified == 1 and "받는분 성명" in result.text


def test_set_one_column_appends_a_new_last_column(table, text):
    grid = build_view(table, FORM)
    changed = set_one_column(grid, None, "보내는분", BLANK_LABEL, "최고다농수산")
    assert len(changed) == len(grid) + 1 and changed.iloc[-1][V_ORDER] == len(grid) + 1
    result = edit(table, text, changed)
    assert result.added == 1
    last = block_of(result).iloc[-1]
    assert last["헤더명"] == "보내는분" and last["고정값(Hardcoded)"] == "최고다농수산"
