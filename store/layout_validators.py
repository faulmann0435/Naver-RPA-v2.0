"""Pure validation of the purchase-order forms (no streamlit): Korean messages, row = 0-based position.

The validators take a layout table in the shape of output_layout.csv (text cells, original Korean
headers; META columns are ignored, rows of other channels are skipped). A per-form problem is
reported on a row of that form (`row` = position in the given frame); problems that belong to no
row (a form that vanishes, a form without columns, the final guard) have `row=None`.
"""
from __future__ import annotations

from collections import Counter
from collections.abc import Collection, Sequence
from dataclasses import dataclass, field

import pandas as pd

from core.dictionary import _parse_bool
from store.csv_codec import from_csv_text
from store.layout_repo import (
    COLUMN,
    FILENAME,
    FIXED,
    FORM,
    HEADER,
    LAYOUT_CHANNEL,
    SOURCE,
    layout_csv_to_raw,
)
from store.rules_validators import ALL_VENDORS, final_guard
from store.validators import Issue

FORBIDDEN_FILENAME_CHARS = '\\/:*?"<>|'
AREA_ROUTE, AREA_OPTIONS, AREA_DICTIONARY = "상품분류(고급 설정)", "옵션규칙(고급 설정)", "품목 사전"


@dataclass(frozen=True)
class LayoutReferences:
    """How often each form name is used by the enabled rows of the rule tables and the dictionary."""

    route: Counter[str] = field(default_factory=Counter)
    options: Counter[str] = field(default_factory=Counter)
    dictionary: Counter[str] = field(default_factory=Counter)


def _s(value: object) -> str:
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()


def _enabled_names(text: str | None, column: str) -> Counter[str]:
    """Names in `column` of the naver rows that are enabled (a missing file counts as no references)."""
    if text is None or not text.strip("﻿ \r\n"):
        return Counter()
    frame = from_csv_text(text, as_text=True)
    if column not in frame.columns:
        return Counter()
    mask = pd.Series(True, index=frame.index)
    if "channel" in frame.columns:
        mask &= frame["channel"] == LAYOUT_CHANNEL
    if "enabled" in frame.columns:
        mask &= frame["enabled"].map(_parse_bool)
    return Counter(name for name in (_s(v) for v in frame.loc[mask, column]) if name)


def collect_references(route_csv: str | None, options_csv: str | None, dictionary_csv: str | None) -> LayoutReferences:
    """Count the form references of product_route.csv (양식명칭), option_rules.csv (양식명칭, not ALL) and dictionary.csv."""
    options = _enabled_names(options_csv, FORM)
    return LayoutReferences(
        route=_enabled_names(route_csv, FORM),
        options=Counter({k: v for k, v in options.items() if k.upper() != ALL_VENDORS}),
        dictionary=_enabled_names(dictionary_csv, "vendor_id"),
    )


# ---------------------------------------------------------------- per-form rows

def _form_rows(layout: pd.DataFrame) -> dict[str, list[tuple[int, dict[str, str]]]]:
    """Rows of the naver channel per form name (first-appearance order); rows without a name under ""."""
    forms: dict[str, list[tuple[int, dict[str, str]]]] = {}
    for position, (_, series) in enumerate(layout.iterrows()):
        values = {str(k): _s(v) for k, v in series.items()}
        if "channel" in values and values["channel"] != LAYOUT_CHANNEL:
            continue
        forms.setdefault(values.get(FORM, ""), []).append((position, values))
    return forms


def _row_label(form: str, values: dict[str, str]) -> str:
    return f"[{form}] 칸 '{values.get(HEADER, '')}'"


def _filename_issues(form: str, rows: list[tuple[int, dict[str, str]]]) -> list[Issue]:
    names = list(dict.fromkeys(values.get(FILENAME, "") for _, values in rows))
    first = rows[0][0]
    issues: list[Issue] = []
    if len(names) > 1:
        shown = ", ".join(f"'{n}'" if n else "(비어 있음)" for n in names)
        issues.append(Issue("error", first, FILENAME, f"[{form}] 파일명이 칸마다 다릅니다 ({shown}). 한 양식에는 파일명이 하나여야 합니다."))
    for name in names:
        bad = sorted({ch for ch in name if ch in FORBIDDEN_FILENAME_CHARS})
        if bad:
            issues.append(Issue(
                "error", first, FILENAME,
                f"[{form}] 파일명 '{name}'에 쓸 수 없는 글자가 있습니다: {' '.join(bad)}  (파일 이름에는 {' '.join(FORBIDDEN_FILENAME_CHARS)} 를 쓸 수 없습니다)",
            ))
    return issues


def _header_issues(form: str, rows: list[tuple[int, dict[str, str]]]) -> list[Issue]:
    counts = Counter(v.get(HEADER, "") for _, v in rows if v.get(HEADER, ""))
    issues: list[Issue] = []
    for position, values in rows:
        header = values.get(HEADER, "")
        if not header:
            issues.append(Issue("error", position, HEADER, f"[{form}] 칸 이름이 비어 있는 칸이 있습니다 (열 {values.get(COLUMN, '?')})."))
        elif counts[header] > 1:
            issues.append(Issue("error", position, HEADER, f"[{form}] 칸 이름 '{header}'이(가) {counts[header]}번 중복되어 있습니다. 칸 이름은 겹치면 안 됩니다."))
    return issues


def _row_warnings(
    form: str, rows: list[tuple[int, dict[str, str]]], known_sources: Collection[str] | None
) -> list[Issue]:
    issues: list[Issue] = []
    for position, values in rows:
        source, fixed = values.get(SOURCE, ""), values.get(FIXED, "")
        if source and fixed:
            issues.append(Issue(
                "warning", position, FIXED,
                f"{_row_label(form, values)}: '넣을 내용'과 '고정 글자'가 둘 다 있습니다. 고정 글자만 쓰이고 '넣을 내용'은 무시됩니다.",
            ))
        if source and known_sources is not None and source not in known_sources:
            issues.append(Issue(
                "warning", position, SOURCE,
                f"{_row_label(form, values)}: 데이터 항목 '{source}'은(는) 알려진 항목이 아닙니다. 주문 파일에 없으면 빈칸으로 나옵니다.",
            ))
    return issues


# ---------------------------------------------------------------- vanished forms

def _where(name: str, refs: LayoutReferences) -> list[str]:
    where: list[str] = []
    if refs.route.get(name):
        where.append(f"{AREA_ROUTE} {refs.route[name]}줄")
    options = sum(v for k, v in refs.options.items() if k.upper() == name.upper())
    if options:
        where.append(f"{AREA_OPTIONS} {options}줄")
    if refs.dictionary.get(name):
        where.append(f"{AREA_DICTIONARY} {refs.dictionary[name]}개")
    return where


def _vanished_issues(old_names: Sequence[str], new_names: Collection[str], refs: LayoutReferences) -> list[Issue]:
    issues: list[Issue] = []
    for name in old_names:
        if name in new_names:
            continue
        where = _where(name, refs)
        if where:
            issues.append(Issue(
                "error", None, FORM,
                f"양식 '{name}'은(는) 아직 쓰이고 있어 없앨 수 없습니다. 쓰는 곳: {', '.join(where)}. "
                "양식 이름을 바꾸는 것도 없애는 것으로 칩니다. 먼저 그 규칙·품목의 양식을 바꾸세요.",
            ))
    return issues


# ---------------------------------------------------------------- entry points

def layout_form_names(layout: pd.DataFrame) -> list[str]:
    """Names of the forms (naver channel, non-empty), in first-appearance order."""
    return [name for name in _form_rows(layout) if name]


def validate_layout(
    layout: pd.DataFrame,
    old_layout: pd.DataFrame | None,
    refs: LayoutReferences,
    known_sources: Collection[str] | None = None,
    empty_forms: Sequence[str] = (),
) -> list[Issue]:
    """Judge the whole layout table.

    old_layout: the table before the edit (forms that vanished while still referenced are errors).
    known_sources: valid 매핑데이터 values (None: do not check). empty_forms: forms the user
    created that have no column yet (error: a form needs at least one column).
    """
    issues: list[Issue] = []
    forms = _form_rows(layout)
    for position, _values in forms.get("", []):
        issues.append(Issue("error", position, FORM, "양식명칭이 비어 있는 칸이 있습니다."))
    for form, rows in forms.items():
        if not form:
            continue
        issues += _filename_issues(form, rows) + _header_issues(form, rows) + _row_warnings(form, rows, known_sources)
    issues += [
        Issue("error", None, FORM, f"양식 '{name}'에 칸이 하나도 없습니다. 칸을 1개 이상 추가하세요.")
        for name in empty_forms if name not in forms
    ]
    if old_layout is not None:
        issues += _vanished_issues(layout_form_names(old_layout), set(forms), refs)
    return issues


def layout_final_guard(route_csv: str, options_csv: str, layout_csv: str) -> list[Issue]:
    """Error when the rules together with this layout text cannot be turned into a working config."""
    try:
        layout = layout_csv_to_raw(layout_csv)
    except (ValueError, KeyError, pd.errors.ParserError) as e:
        return [Issue("error", None, "", f"양식 파일을 읽을 수 없습니다. ({type(e).__name__}: {str(e)[:200]})")]
    return final_guard(route_csv, options_csv, layout)
