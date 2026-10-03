"""Pure helpers of the product page (no streamlit): grouping, dirty marks, applying edits, validation."""
from __future__ import annotations

import hashlib
from collections import Counter
from collections.abc import Collection, Mapping
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone

import pandas as pd

from core.dictionary import DictionarySettings, _parse_bool
from store.dictionary_repo import (
    RowKey,
    group_id,
    group_id_of,
    group_ids,
    now_kst_iso,
    product_rows,
    text_records,
)
from store.rules_repo import DICTIONARY_COLUMNS
from store.validators import validate_dictionary
from ui.dictionary_logic import ALL_VENDORS, describe_issue, format_created
from ui.option_logic import FormValues, row_fields, values_from_row

COL_NAME, COL_OPTIONS, COL_REVIEW, COL_DIRTY = "상품명", "옵션 수", "확인 필요", "저장 안 됨"
COL_CREATED = "등록일"
__all__ = ["group_id", "group_id_of"]
DIRTY_MARK = "✎"


@dataclass(frozen=True)
class _Delete:
    """Marker: delete this option from the dictionary."""


DELETE = _Delete()
EditValue = FormValues | _Delete


@dataclass(frozen=True)
class ProductSummary:
    group: str
    product_no: str
    name: str
    option_count: int
    review_count: int
    created_at: str = ""  # earliest non-empty created_at of the options (ISO text; empty if none)


def _text(frame: pd.DataFrame) -> pd.DataFrame:
    return frame.reindex(columns=DICTIONARY_COLUMNS, fill_value="").fillna("").astype(str).reset_index(drop=True)


def option_prefix(product_no: str, option_key: str) -> str:
    """Stable widget-key prefix of one option (the option text is hashed: it may hold any character)."""
    digest = hashlib.sha1(option_key.encode("utf-8")).hexdigest()[:10]
    return f"pm_{product_no}_{digest}"


def group_key(group: str) -> str:
    """Widget-key part of a group (hashed: the group id holds a separator and the product name)."""
    return hashlib.sha1(group.encode("utf-8")).hexdigest()[:12]


def product_name(rows: list[dict[str, str]]) -> str:
    """The latest product name of a group's rows (by last_seen_at, later rows win ties; empty -> product_no)."""
    named = [(r.get("last_seen_at", ""), i, r["product_name_ref"].strip()) for i, r in enumerate(rows)
             if r["product_name_ref"].strip()]
    return max(named)[2] if named else (rows[0]["product_no"] if rows else "")


def _matches(group: pd.DataFrame, name: str, query: str, vendor: str) -> bool:
    if vendor and vendor != ALL_VENDORS and not (group["vendor_id"] == vendor).any():
        return False
    text = query.strip().lower()
    if not text:
        return True
    haystack = " ".join(
        [name, *group["option_raw_ref"], *group["option_key"], *group["display_template"], *group["product_no"]]
    ).lower()
    return text in haystack


def product_summaries(
    frame: pd.DataFrame, query: str = "", vendor: str = ALL_VENDORS, review_only: bool = False
) -> list[ProductSummary]:
    """Groups (product_no + product name) with the latest name; unchecked ones first, then by name."""
    text = _text(frame)
    summaries: list[ProductSummary] = []
    for gid, group in text.groupby(group_ids(text), sort=False):
        name = product_name(text_records(group))
        review = int(group["needs_review"].map(_parse_bool).sum())
        if review_only and review == 0:
            continue
        if _matches(group, name, query, vendor):
            summaries.append(ProductSummary(
                str(gid), str(group["product_no"].iloc[0]), name, len(group), review, earliest_created(group)
            ))
    return sorted(summaries, key=lambda s: (s.review_count == 0, s.name))


def earliest_created(group: pd.DataFrame) -> str:
    """The earliest readable created_at among a group's options ('' when none)."""
    stamps = [(moment, text) for text in group["created_at"] if (moment := _moment(text)) is not None]
    return min(stamps)[1] if stamps else ""


def _moment(text: str) -> datetime | None:
    try:
        moment = datetime.fromisoformat(text.strip())
    except ValueError:
        return None
    return moment if moment.tzinfo else moment.replace(tzinfo=timezone(timedelta(hours=9)))


def summary_table(summaries: list[ProductSummary], dirty: Collection[str]) -> pd.DataFrame:
    return pd.DataFrame(
        {
            COL_REVIEW: [s.review_count for s in summaries],
            COL_OPTIONS: [s.option_count for s in summaries],
            COL_DIRTY: [DIRTY_MARK if s.group in dirty else "" for s in summaries],
            COL_NAME: [s.name for s in summaries],
            COL_CREATED: [format_created(s.created_at) for s in summaries],
        },
        columns=[COL_REVIEW, COL_OPTIONS, COL_DIRTY, COL_NAME, COL_CREATED],
    )


def product_option_rows(frame: pd.DataFrame, group: str) -> list[dict[str, str]]:
    """The group's rows (text) in file order."""
    return [r for r in text_records(frame) if group_id_of(r) == group]


def common_vendor(rows: list[dict[str, str]]) -> str:
    """The most frequent 발주양식 among the options (first one wins a tie)."""
    counts = Counter(r["vendor_id"] for r in rows)
    return counts.most_common(1)[0][0] if counts else ""


def dirty_products(
    frame: pd.DataFrame, pending: Mapping[str, FormValues], deleted: Collection[str]
) -> set[str]:
    """Group ids having a pending edit that differs from the saved row, or a delete mark."""
    dirty: set[str] = set()
    for row in text_records(frame):
        prefix = option_prefix(row["product_no"], row["option_key"])
        edit = pending.get(prefix)
        if prefix in deleted or (edit is not None and edit != values_from_row(row)):
            dirty.add(group_id_of(row))
    return dirty


def apply_product_edits(
    base: pd.DataFrame,
    group: str,
    values_by_key: Mapping[RowKey, EditValue],
    user: str = "",
    now: str | None = None,
) -> pd.DataFrame:
    """New frame: this group's options replaced / deleted by key; changed rows get updated_at/by."""
    stamp = now or now_kst_iso()
    rows: list[dict[str, str]] = []
    for rec in text_records(base):
        edit = values_by_key.get((rec["product_no"], rec["option_key"])) if group_id_of(rec) == group else None
        if edit is None:
            rows.append(rec)
        elif isinstance(edit, FormValues):
            fields = row_fields(edit)
            changed = fields != row_fields(values_from_row(rec))
            rows.append({**rec, **fields, "updated_at": stamp, "updated_by": user} if changed else rec)
    return pd.DataFrame(rows, columns=DICTIONARY_COLUMNS)


def change_counts(base: pd.DataFrame, new: pd.DataFrame, group: str) -> tuple[int, int]:
    """(modified, deleted) option rows of one group between two frames."""
    before, after = product_rows(base, group), product_rows(new, group)
    deleted = sum(1 for key in before if key not in after)
    modified = sum(1 for key, row in after.items() if key in before and before[key] != row)
    return modified, deleted


def product_message(name: str, modified: int, deleted: int) -> str:
    return f"품목 수정: {name} (수정 {modified}, 삭제 {deleted})"


def product_issues(
    frame: pd.DataFrame, group: str, vendor_ids: list[str], settings: DictionarySettings
) -> tuple[list[str], list[str]]:
    """(errors, warnings) of THIS group's rows; problems elsewhere in the file are ignored."""
    text = _text(frame)
    ids = group_ids(text)
    errors: list[str] = []
    warnings: list[str] = []
    for issue in validate_dictionary(text, vendor_ids, settings):
        if issue.row is not None and ids.at[issue.row] != group:
            continue
        (errors if issue.level == "error" else warnings).append(describe_issue(text, issue))
    return errors, warnings
