"""History of output_layout.csv: readable diff and revert guard (offline)."""
from collections import Counter

import pytest

from core.config_loader import read_config_sheets
from store.layout_repo import OUTPUT_LAYOUT_FILE, layout_sheet_to_csv
from store.layout_validators import LayoutReferences
from tests.regression_harness import CONFIG_PATH
from ui.history_logic import (
    LAYOUT_DIFF_COLUMNS,
    TARGETS,
    diff_against_current,
    diff_layout,
    file_label,
    layout_revert_errors,
)

ROUTE = "우선순위,키워드,양식명칭\n1,DEFAULT,메로 발주양식\n"
OPTIONS = "순서,적용대상(상품명),ActionType (명령),양식명칭,Parameter (설정값)\n1,ALL,REMOVE_TEXT,ALL,x\n"
FORM = "메로 발주양식"


@pytest.fixture(scope="module")
def text() -> str:
    sheet = read_config_sheets(str(CONFIG_PATH))["OutputLayout"]
    return layout_sheet_to_csv(sheet, "migration", "2026-10-02T00:00:00+09:00")


def test_target_is_registered():
    assert TARGETS["발주서 양식"] == OUTPUT_LAYOUT_FILE and file_label(OUTPUT_LAYOUT_FILE) == "발주서 양식"


def test_identical_versions_have_no_diff(text):
    assert diff_layout(text, text).empty


def test_diff_reports_rename_source_fixed_added_and_removed(text):
    now = text.replace(f"{FORM},{FORM},A,수취인명,수취인명,", f"{FORM},{FORM},A,수취인명,수취인명,고정값")
    now = now.replace("메로 발주양식,메로 발주양식,K,배송메세지,배송메세지,", "메로 발주양식,메로 발주양식,K,배송메모,배송메세지,")
    now += "메로 발주양식,메로 발주양식,L,새칸,,,naver,,\n"
    diff = diff_layout(text, now)
    assert list(diff.table.columns) == LAYOUT_DIFF_COLUMNS
    by_name = {row["칸 이름"]: row for _, row in diff.table.iterrows()}
    assert by_name["수취인명"]["구분"] == "변경" and "고정 글자" in by_name["수취인명"]["변경 내용"]
    assert by_name["새칸"]["구분"] == "추가" and "되돌리면 사라짐" in by_name["새칸"]["변경 내용"]
    assert by_name["배송메세지"]["구분"] == "삭제" and "되돌리면 다시 생김" in by_name["배송메세지"]["변경 내용"]
    assert by_name["배송메모"]["구분"] == "추가"
    assert set(diff.table["양식"]) == {FORM}


def test_diff_order_is_relative_among_common_columns(text):
    lines = text.splitlines()
    index = next(i for i, ln in enumerate(lines) if ln.startswith(f"{FORM},{FORM},B,"))
    removed = "\n".join(lines[:index] + lines[index + 1:]) + "\n"  # one column deleted, no letters renumbered
    assert [r["구분"] for _, r in diff_layout(text, removed).table.iterrows()] == ["삭제"]


def test_diff_moved_column_is_reported(text):
    swapped = text.replace(f"{FORM},{FORM},A,수취인명", f"{FORM},{FORM},Z,수취인명")
    diff = diff_layout(text, swapped).table
    assert "순서가 바뀜" in " ".join(diff["변경 내용"])


def test_diff_against_current_dispatches_to_layout(text):
    assert diff_against_current(OUTPUT_LAYOUT_FILE, text, text, None).empty  # type: ignore[arg-type]


def test_revert_errors_for_referenced_vanishing_form(text):
    without = "\n".join(ln for ln in text.splitlines() if FORM not in ln) + "\n"
    refs = LayoutReferences(route=Counter({FORM: 1}))
    errors = layout_revert_errors(without, text, ROUTE, OPTIONS, refs)
    assert any("아직 쓰이고 있어" in e for e in errors)
    assert layout_revert_errors(text, without, ROUTE, OPTIONS, refs) == []  # restoring a form is fine


def test_revert_errors_for_invalid_old_layout(text):
    broken = text.replace(f"{FORM},{FORM},B,공란1", f"{FORM},{FORM},B,수취인명")
    assert any("중복" in e for e in layout_revert_errors(broken, text, ROUTE, OPTIONS, LayoutReferences()))
    header_only = text.splitlines()[0] + "\n"
    assert any("처리할 수 없습니다" in e for e in layout_revert_errors(header_only, text, ROUTE, OPTIONS, LayoutReferences()))
