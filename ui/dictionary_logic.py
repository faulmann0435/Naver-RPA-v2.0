"""Pure helpers of the dictionary page (no streamlit): filtered views, edit folding, diffs, save preparation."""
from __future__ import annotations

import io
from collections.abc import Sequence

import pandas as pd

from core.dictionary import (
    DictionaryEntry,
    DictionarySettings,
    ItemDictionary,
    _parse_bool,
    _parse_float,
    render_entry,
)
from core.merger import _format_weight
from core.template import format_number
from store.dictionary_repo import now_kst_iso
from store.rules_repo import DICTIONARY_COLUMNS
from store.validators import COLUMN_LABELS, Issue

RID = "_rid"
DELETE = "_delete"
V_ENABLED, V_REVIEW, V_NAME, V_OPTION = "사용", "검토 필요", "상품명", "옵션"
V_VENDOR, V_TEMPLATE, V_QTY1 = "발주양식", "발주서 표기", "수량 1일 때 표기"
V_GROUP, V_WEIGHT, V_APPEND, V_DELETE = "합산 이름", "1개당 무게(kg)", "묶음 끝에 붙임", "삭제"
BASIC_VIEW = [V_ENABLED, V_REVIEW, V_NAME, V_OPTION, V_VENDOR, V_TEMPLATE, V_QTY1, V_DELETE]
ADVANCED_VIEW = [V_ENABLED, V_REVIEW, V_NAME, V_OPTION, V_VENDOR, V_TEMPLATE, V_QTY1, V_GROUP, V_WEIGHT, V_APPEND, V_DELETE]
DISABLED_VIEW = [V_NAME, V_OPTION]
ALL_VENDORS = "전체"
EXPORT_HEADERS = {
    "channel": "채널", "product_no": "상품번호", "option_key": "옵션(정규화)", "product_name_ref": "상품명",
    "option_raw_ref": "옵션", "vendor_id": "발주양식", "display_template": "발주서 표기",
    "display_template_qty1": "수량 1일 때 표기", "sum_group": "합산 이름", "unit_weight_kg": "1개당 무게(kg)",
    "append_to_end": "묶음 끝에 붙임", "needs_review": "검토 필요", "enabled": "사용", "source": "출처",
    "last_seen_at": "마지막 주문 확인", "updated_at": "수정 시각", "updated_by": "수정한 사람",
}
# dictionary column -> (view column, kind)
_BOOL_FIELDS = {"enabled": V_ENABLED, "needs_review": V_REVIEW, "append_to_end": V_APPEND}
_TEXT_FIELDS = {
    "vendor_id": V_VENDOR, "display_template": V_TEMPLATE, "display_template_qty1": V_QTY1, "sum_group": V_GROUP,
}
EDIT_FIELDS = [*_BOOL_FIELDS, *_TEXT_FIELDS, "unit_weight_kg"]
FIELD_LABELS = {
    "enabled": "사용", "needs_review": "검토 필요", "append_to_end": "묶음 끝에 붙임", "vendor_id": "발주양식",
    "display_template": "발주서 표기", "display_template_qty1": "수량 1일 때 표기", "sum_group": "합산 이름",
    "unit_weight_kg": "1개당 무게(kg)",
}


def _s(value: object) -> str:
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()


def make_work(frame: pd.DataFrame) -> pd.DataFrame:
    """Working copy: all DICTIONARY_COLUMNS as text, a `_delete` flag, index = row id (position in the base)."""
    work = frame.reindex(columns=DICTIONARY_COLUMNS, fill_value="").fillna("").astype(str).reset_index(drop=True)
    work[DELETE] = False
    return work


def filter_rids(
    work: pd.DataFrame, query: str, vendor: str, review_only: bool
) -> list[int]:
    """Row ids (positions) matching the filters; the search looks at name, option and display text."""
    mask = pd.Series(True, index=work.index)
    text = query.strip().lower()
    if text:
        haystack = (
            work["product_name_ref"] + " " + work["option_raw_ref"] + " " + work["option_key"]
            + " " + work["display_template"] + " " + work["display_template_qty1"]
        ).str.lower()
        mask &= haystack.str.contains(text, regex=False)
    if vendor and vendor != ALL_VENDORS:
        mask &= work["vendor_id"] == vendor
    if review_only:
        mask &= work["needs_review"].map(_parse_bool)
    return [int(i) for i in work.index[mask]]


def build_view(work: pd.DataFrame, rids: Sequence[int], advanced: bool) -> pd.DataFrame:
    """The rows shown in the editor; index and `_rid` both hold the row id."""
    part = work.loc[list(rids)]
    view = pd.DataFrame(index=part.index)
    view[RID] = part.index.astype(int)
    view[V_ENABLED] = part["enabled"].map(_parse_bool)
    view[V_REVIEW] = part["needs_review"].map(_parse_bool)
    view[V_NAME] = part["product_name_ref"]
    view[V_OPTION] = part["option_raw_ref"]
    view[V_VENDOR] = part["vendor_id"]
    view[V_TEMPLATE] = part["display_template"]
    view[V_QTY1] = part["display_template_qty1"]
    view[V_GROUP] = part["sum_group"]
    view[V_WEIGHT] = pd.to_numeric(part["unit_weight_kg"], errors="coerce").astype(float)
    view[V_APPEND] = part["append_to_end"].map(_parse_bool)
    view[V_DELETE] = part[DELETE].astype(bool)
    return view[[RID, *(ADVANCED_VIEW if advanced else BASIC_VIEW)]]


def _bool_text(flag: bool) -> str:
    return "1" if flag else "0"


def _weight_text(value: object) -> str:
    number = _parse_float(value) if not isinstance(value, (int, float)) else float(value)
    return "" if number is None or pd.isna(number) else format_number(number)


def _apply_cell(work: pd.DataFrame, rid: int, column: str, value: object) -> None:
    if column in _BOOL_FIELDS:
        new = _parse_bool(value)
        if _parse_bool(work.at[rid, column]) != new:
            work.at[rid, column] = _bool_text(new)
    elif column == "unit_weight_kg":
        new_text = _weight_text(value)
        if _weight_text(work.at[rid, column]) != new_text:
            work.at[rid, column] = new_text
    elif _s(work.at[rid, column]) != _s(value):
        work.at[rid, column] = _s(value)


def fold_edits(work: pd.DataFrame, edited: pd.DataFrame) -> pd.DataFrame:
    """New working copy with the editor's rows folded in by `_rid`; rows not in `edited` are untouched."""
    updated = work.copy()
    columns = {**_BOOL_FIELDS, **_TEXT_FIELDS, "unit_weight_kg": V_WEIGHT}
    for _, row in edited.iterrows():
        rid = int(row[RID])
        if rid not in updated.index:
            continue
        for column, view_column in columns.items():
            if view_column in edited.columns:
                _apply_cell(updated, rid, column, row[view_column])
        if V_DELETE in edited.columns:
            updated.at[rid, DELETE] = bool(row[V_DELETE])
    return updated


BULK_FIELDS = {V_ENABLED: "enabled", V_REVIEW: "needs_review", V_DELETE: DELETE}


def bulk_set(work: pd.DataFrame, rids: Sequence[int], label: str, value: bool) -> pd.DataFrame:
    """New working copy with one checkbox column (사용 / 검토 필요 / 삭제) set for the given row ids."""
    column = BULK_FIELDS[label]
    updated = work.copy()
    targets = [rid for rid in rids if rid in updated.index]
    updated.loc[targets, column] = value if column == DELETE else _bool_text(value)
    return updated


def _differences(before: pd.Series, after: pd.Series) -> list[str]:
    changes = []
    for column in EDIT_FIELDS:
        old, new = _s(before[column]), _s(after[column])
        if column in _BOOL_FIELDS:
            old, new = _bool_text(_parse_bool(old)), _bool_text(_parse_bool(new))
        elif column == "unit_weight_kg":
            old, new = _weight_text(old), _weight_text(new)
        if old != new:
            changes.append(f"{FIELD_LABELS[column]}: {old or '(비움)'} → {new or '(비움)'}")
    return changes


def changed_rids(base: pd.DataFrame, work: pd.DataFrame) -> tuple[list[int], list[int]]:
    """(modified row ids, deleted row ids); a deleted row is not counted as modified."""
    deleted = [int(i) for i in work.index[work[DELETE]]]
    modified = [
        int(i) for i in work.index
        if i not in deleted and i in base.index and _differences(base.loc[i], work.loc[i])
    ]
    return modified, deleted


def example_text(row: pd.Series, qty: int) -> str:
    """Display text of a dictionary row for `qty` units."""
    group = _s(row["sum_group"])
    weight = _parse_float(row["unit_weight_kg"])
    if group and weight:
        return f"{group} {_format_weight(weight * qty)}"
    entry = DictionaryEntry(
        product_no=_s(row["product_no"]), option_key=_s(row["option_key"]), vendor_id=_s(row["vendor_id"]),
        display_template=_s(row["display_template"]), display_template_qty1=_s(row["display_template_qty1"]),
    )
    try:
        return render_entry(entry, qty)
    except ValueError:
        return "(표기 오류)"


def change_table(base: pd.DataFrame, work: pd.DataFrame) -> pd.DataFrame:
    """Changed rows: before -> after of the changed fields, plus the qty 2 rendering."""
    modified, deleted = changed_rids(base, work)
    records = [
        {
            "구분": "수정", "상품명": work.at[rid, "product_name_ref"], "옵션": work.at[rid, "option_raw_ref"],
            "변경 내용": " / ".join(_differences(base.iloc[rid], work.iloc[rid])),
            "수량 2 결과": example_text(work.iloc[rid], 2),
        }
        for rid in modified
    ] + [
        {
            "구분": "삭제", "상품명": work.at[rid, "product_name_ref"], "옵션": work.at[rid, "option_raw_ref"],
            "변경 내용": "사전에서 삭제", "수량 2 결과": "",
        }
        for rid in deleted
    ]
    return pd.DataFrame(records, columns=["구분", "상품명", "옵션", "변경 내용", "수량 2 결과"])


def kept_rows(work: pd.DataFrame) -> pd.DataFrame:
    """Rows that will be saved (not marked for deletion); the original row ids stay as the index."""
    return work[~work[DELETE]][DICTIONARY_COLUMNS]


def describe_issue(kept: pd.DataFrame, issue: Issue) -> str:
    """One Korean line for an issue: which row (product / option) and column, then the message."""
    if issue.row is None:
        return issue.message
    row = kept.iloc[issue.row]
    name = _s(row["product_name_ref"]) or _s(row["product_no"])
    option = _s(row["option_raw_ref"]) or _s(row["option_key"]) or "옵션 없음"
    column = COLUMN_LABELS.get(issue.column, issue.column)
    return f"[{name} / {option}] {column}: {issue.message}"


def prepare_save(base: pd.DataFrame, work: pd.DataFrame, user: str, now: str | None = None) -> pd.DataFrame:
    """The frame to write: deleted rows dropped, updated_at/updated_by set on modified rows."""
    modified, _deleted = changed_rids(base, work)
    stamp = now or now_kst_iso()
    out = work.copy()
    out.loc[modified, "updated_at"] = stamp
    out.loc[modified, "updated_by"] = user
    return kept_rows(out).reset_index(drop=True)


def dictionary_from_work(work: pd.DataFrame, settings: DictionarySettings) -> ItemDictionary:
    """ItemDictionary built from an (unsaved) working copy, deleted rows excluded."""
    return ItemDictionary(kept_rows(work), settings)


def to_excel_bytes(frame: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    frame.reindex(columns=DICTIONARY_COLUMNS, fill_value="").rename(columns=EXPORT_HEADERS).to_excel(
        buffer, index=False, sheet_name="품목사전"
    )
    return buffer.getvalue()
