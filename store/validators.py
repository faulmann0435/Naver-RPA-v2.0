"""Pure validation of item-dictionary rows (no streamlit): Korean messages, row = 0-based position."""
from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Literal

import pandas as pd

from core.dictionary import DictionarySettings, _parse_bool, _parse_float, _text
from core.option_key import make_option_key, normalize_product_no
from core.template import has_qty_placeholder, validate_template

COLUMN_LABELS: dict[str, str] = {
    "product_no": "상품번호",
    "option_key": "옵션",
    "vendor_id": "발주양식",
    "display_template": "발주서 표기",
    "display_template_qty1": "수량 1일 때 표기",
    "sum_group": "합산 이름",
    "unit_weight_kg": "1개당 무게(kg)",
    "append_to_end": "묶음 끝에 붙임",
}
NO_QTY_WARNING = "수량 칸이 없습니다. 2개 이상 주문 시 (xN)이 자동으로 붙습니다"
_PLACEHOLDER = re.compile(r"\{[^{}]*\}")


@dataclass(frozen=True)
class Issue:
    level: Literal["error", "warning"]
    row: int | None
    column: str
    message: str


def _error(row: int | None, column: str, message: str) -> Issue:
    return Issue("error", row, column, message)


def _template_issues(row: int | None, column: str, template: str) -> list[Issue]:
    return [_error(row, column, message) for message in validate_template(template)]


def _vendor_issues(row: int | None, vendor_id: str, vendor_ids: Sequence[str]) -> list[Issue]:
    if vendor_id in vendor_ids:
        return []
    if not vendor_id:
        return [_error(row, "vendor_id", "발주양식이 비어 있습니다.")]
    return [_error(row, "vendor_id", f"발주양식 '{vendor_id}'을(를) 찾을 수 없습니다.")]


def _weight_issues(row: int | None, sum_group: str, raw_weight: str) -> list[Issue]:
    weight = _parse_float(raw_weight)
    if sum_group:
        if weight is None or weight <= 0:
            return [_error(row, "unit_weight_kg", "합산 이름이 있으면 1개당 무게(kg)가 0보다 커야 합니다.")]
        return []
    if raw_weight:
        return [_error(row, "unit_weight_kg", "1개당 무게(kg)를 쓰려면 합산 이름도 필요합니다.")]
    return []


def _row_issues(row: int | None, values: Mapping[str, object], vendor_ids: Sequence[str]) -> list[Issue]:
    template = _text(values.get("display_template"))
    qty1_template = _text(values.get("display_template_qty1"))
    sum_group = _text(values.get("sum_group"))
    issues: list[Issue] = []
    if not normalize_product_no(values.get("product_no")):
        issues.append(_error(row, "product_no", "상품번호가 비어 있습니다."))
    issues += _vendor_issues(row, _text(values.get("vendor_id")), vendor_ids)
    issues += _template_issues(row, "display_template", template)
    issues += _template_issues(row, "display_template_qty1", qty1_template)
    if not template and not sum_group:
        issues.append(_error(row, "display_template", "발주서 표기와 합산 이름이 모두 비어 있습니다."))
    issues += _weight_issues(row, sum_group, _text(values.get("unit_weight_kg")))
    if sum_group and _parse_bool(values.get("append_to_end")):
        issues.append(_error(row, "append_to_end", "합산 이름과 '묶음 끝에 붙임'은 함께 쓸 수 없습니다."))
    if template and not sum_group and not validate_template(template) and not has_qty_placeholder(template):
        issues.append(Issue("warning", row, "display_template", NO_QTY_WARNING))
    return issues


def _entry_key(values: Mapping[str, object], settings: DictionarySettings) -> tuple[str, str]:
    return (
        normalize_product_no(values.get("product_no")),
        make_option_key(_text(values.get("option_key")), settings.ignored_option_groups),
    )


def _duplicate_issues(
    keys: Mapping[int, tuple[str, str]],
) -> list[Issue]:
    positions: dict[tuple[str, str], list[int]] = {}
    for row, key in keys.items():
        if key[0]:
            positions.setdefault(key, []).append(row)
    issues: list[Issue] = []
    for rows in positions.values():
        if len(rows) < 2:
            continue
        where = ", ".join(str(r + 1) for r in rows)
        issues += [
            _error(r, "option_key", f"같은 상품번호·옵션이 {len(rows)}곳에 중복되어 있습니다 (행 {where}).")
            for r in rows
        ]
    return issues


def validate_dictionary(
    df: pd.DataFrame, vendor_ids: Sequence[str], settings: DictionarySettings
) -> list[Issue]:
    """Validate the enabled rows of a dictionary frame; `row` is the position within `df`."""
    issues: list[Issue] = []
    keys: dict[int, tuple[str, str]] = {}
    for position, (_, series) in enumerate(df.iterrows()):
        values: dict[str, object] = {str(k): v for k, v in series.items()}
        if "enabled" in values and not _parse_bool(values["enabled"]):
            continue
        issues += _row_issues(position, values, vendor_ids)
        keys[position] = _entry_key(values, settings)
    return issues + _duplicate_issues(keys)


def _standalone_number(template: str, number: int) -> bool:
    text = _PLACEHOLDER.sub("", template)
    return re.search(rf"(?<![\d.]){number}(?!\.?\d)", text) is not None


def validate_new_entry(
    row: Mapping[str, object],
    qty_example: int,
    vendor_ids: Sequence[str],
    settings: DictionarySettings,
) -> list[Issue]:
    """Same rules as validate_dictionary for one new row, plus a check for a hard-coded example quantity."""
    del settings  # reserved: key normalization is done by the caller / append_entries
    issues = _row_issues(None, row, vendor_ids)
    template = _text(row.get("display_template"))
    if qty_example > 1 and _standalone_number(template, qty_example):
        issues.append(
            Issue(
                "warning", None, "display_template",
                f"등록 당시 수량({qty_example})이 표기에 그대로 있습니다. {{수량}}으로 바꿔야 하지 않나요?",
            )
        )
    return issues
