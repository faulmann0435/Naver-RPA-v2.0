"""Pure logic of the dictionary Excel import (no streamlit): read the export format, plan the changes."""
from __future__ import annotations

import io
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass

import pandas as pd

from core.dictionary import DictionarySettings, _parse_bool
from store.dictionary_repo import entry_key
from store.rules_repo import DICTIONARY_COLUMNS
from store.validators import Issue, validate_dictionary
from ui.dictionary_logic import (
    EDIT_FIELDS,
    EXPORT_HEADERS,
    _bool_text,
    apply_cell,
    describe_issue,
    field_differences,
)

SOURCE_EXCEL = "excel"
CHANGE_COLUMNS = ["구분", "상품명", "옵션", "변경 내용"]
_BY_HEADER = {header: column for column, header in EXPORT_HEADERS.items()}
RowKey = tuple[str, str]


class ImportFormatError(ValueError):
    """The uploaded file is not a dictionary export (message is Korean, for the user)."""


@dataclass(frozen=True)
class ImportPlan:
    frame: pd.DataFrame  # the whole dictionary after the import (text, DICTIONARY_COLUMNS)
    touched: list[int]  # positions in `frame` of added and modified rows
    added: int
    modified: int
    unchanged: int
    deleted: int
    changes: pd.DataFrame  # CHANGE_COLUMNS (추가 / 수정 / 삭제)
    errors: list[str]  # problems of the file itself (e.g. the same item twice)


@dataclass(frozen=True)
class ImportJudgement:
    errors: list[str]
    warnings: list[str]
    reference: list[str]


def _s(value: object) -> str:
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()


def read_import_frame(data: bytes) -> pd.DataFrame:
    """Rows of an exported .xlsx as text with DICTIONARY_COLUMNS; unknown / missing columns are an error."""
    try:
        raw = pd.read_excel(io.BytesIO(data), dtype=str, keep_default_na=False)
    except (ValueError, OSError, KeyError, zipfile.BadZipFile) as e:
        raise ImportFormatError(f"엑셀 파일을 읽을 수 없습니다. .xlsx 파일인지 확인하세요. ({type(e).__name__})") from e
    headers = [str(c).strip() for c in raw.columns]
    unknown = [h for h in headers if h not in _BY_HEADER]
    missing = [h for h in EXPORT_HEADERS.values() if h not in headers]
    if unknown or missing:
        parts = []
        if unknown:
            parts.append(f"알 수 없는 열: {', '.join(unknown)}")
        if missing:
            parts.append(f"빠진 열: {', '.join(missing)}")
        raise ImportFormatError(
            "품목 사전에서 내보낸 엑셀 형식이 아닙니다. " + " / ".join(parts)
            + ". '엑셀로 내보내기'로 받은 파일을 고쳐서 올려 주세요. (열 이름과 개수를 바꾸지 마세요)"
        )
    raw.columns = pd.Index([_BY_HEADER[h] for h in headers])
    frame = raw.reindex(columns=DICTIONARY_COLUMNS, fill_value="").fillna("").astype(str)
    frame = frame.apply(lambda col: col.str.strip())
    return frame[(frame != "").any(axis=1)].reset_index(drop=True)


def _text_frame(frame: pd.DataFrame) -> pd.DataFrame:
    return frame.reindex(columns=DICTIONARY_COLUMNS, fill_value="").fillna("").astype(str).reset_index(drop=True)


def _keys(frame: pd.DataFrame, settings: DictionarySettings) -> list[RowKey]:
    return [entry_key(r["product_no"], r["option_key"], settings) for _, r in frame.iterrows()]


def _first_positions(keys: Sequence[RowKey]) -> dict[RowKey, int]:
    positions: dict[RowKey, int] = {}
    for i, key in enumerate(keys):
        positions.setdefault(key, i)
    return positions


def _duplicate_errors(keys: Sequence[RowKey], incoming: pd.DataFrame) -> list[str]:
    seen: dict[RowKey, list[int]] = {}
    for i, key in enumerate(keys):
        seen.setdefault(key, []).append(i)
    errors = []
    for rows in seen.values():
        if len(rows) > 1:
            first = incoming.iloc[rows[0]]
            where = ", ".join(str(r + 2) for r in rows)  # +2: header row, 1-based
            errors.append(
                f"엑셀에 같은 상품번호·옵션이 중복되어 있습니다 (엑셀 {where}번째 줄): "
                f"{_s(first['product_name_ref']) or _s(first['product_no'])} / {_s(first['option_raw_ref']) or '옵션 없음'}"
            )
    return errors


def _new_entry_row(row: pd.Series, key: RowKey, user: str, now: str) -> dict[str, str]:
    out = {column: _s(row[column]) for column in DICTIONARY_COLUMNS}
    out["product_no"], out["option_key"] = key
    out["channel"] = out["channel"] or "naver"
    out["enabled"] = _bool_text(_parse_bool(out["enabled"] or "1"))
    out["needs_review"] = _bool_text(_parse_bool(out["needs_review"]))
    out["append_to_end"] = _bool_text(_parse_bool(out["append_to_end"]))
    out["source"] = SOURCE_EXCEL
    out["updated_at"], out["updated_by"] = now, user
    return out


def _updated_row(current: pd.Series, incoming: pd.Series) -> pd.Series:
    one = pd.DataFrame([current], index=[0])
    for column in EDIT_FIELDS:
        apply_cell(one, 0, column, incoming[column])
    return one.iloc[0]


def _name(row: pd.Series) -> tuple[str, str]:
    return (
        _s(row["product_name_ref"]) or _s(row["product_no"]),
        _s(row["option_raw_ref"]) or _s(row["option_key"]) or "옵션 없음",
    )


def plan_import(
    current: pd.DataFrame, incoming: pd.DataFrame, settings: DictionarySettings,
    user: str, now: str, delete_missing: bool = False,
) -> ImportPlan:
    """Compare by (product_no, normalized option_key). Only the editable fields of existing rows change;
    new rows are added with source 'excel'; rows missing from the file are removed only on request."""
    cur, inc = _text_frame(current), _text_frame(incoming)
    inc_keys = _keys(inc, settings)
    inc_pos = _first_positions(inc_keys)
    errors = _duplicate_errors(inc_keys, inc)
    cur_keys = _keys(cur, settings)
    rows: list[pd.Series] = []
    touched: list[int] = []
    changes: list[list[str]] = []
    unchanged = modified = deleted = 0
    seen: set[RowKey] = set()
    cur_set = set(cur_keys)
    for i, key in enumerate(cur_keys):
        row = cur.iloc[i]
        if key not in inc_pos or key in seen:
            if key in inc_pos or not delete_missing:
                rows.append(row)
            else:
                deleted += 1
                changes.append(["삭제", *_name(row), "사전에서 삭제"])
            continue
        seen.add(key)
        after = _updated_row(row, inc.iloc[inc_pos[key]])
        diffs = field_differences(row, after)
        if not diffs:
            unchanged += 1
            rows.append(row)
            continue
        after = after.copy()
        after["updated_at"], after["updated_by"] = now, user
        touched.append(len(rows))
        rows.append(after)
        modified += 1
        changes.append(["수정", *_name(row), " / ".join(diffs)])
    added = 0
    for key, pos in inc_pos.items():
        if key in cur_set:
            continue
        entry = pd.Series(_new_entry_row(inc.iloc[pos], key, user, now))
        touched.append(len(rows))
        rows.append(entry)
        added += 1
        changes.append(["추가", *_name(entry), "새 항목 (출처: 엑셀)"])
    frame = pd.DataFrame([r.to_dict() for r in rows], columns=DICTIONARY_COLUMNS)
    return ImportPlan(frame, touched, added, modified, unchanged, deleted, pd.DataFrame(changes, columns=CHANGE_COLUMNS), errors)


def judge_import(plan: ImportPlan, vendor_ids: Sequence[str], settings: DictionarySettings) -> ImportJudgement:
    """Validate the whole resulting dictionary; problems in added / modified rows are the ones that count."""
    issues: list[Issue] = validate_dictionary(plan.frame, vendor_ids, settings)
    touched = set(plan.touched)
    errors, warnings, reference = list(plan.errors), [], []
    for issue in issues:
        line = describe_issue(plan.frame, issue)
        if issue.row is not None and issue.row not in touched:
            reference.append(line)
        elif issue.level == "error":
            errors.append(line)
        else:
            warnings.append(line)
    return ImportJudgement(errors, warnings, reference)


def import_message(plan: ImportPlan) -> str:
    return f"엑셀 가져오기: 추가 {plan.added}, 수정 {plan.modified}, 삭제 {plan.deleted}"

