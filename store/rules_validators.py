"""Pure validation of the rule tables (no streamlit): Korean messages, row = 0-based position.

Both validators take a "canonical" frame (English column names, text cells):
  route:   priority, keyword, vendor, enabled
  options: order, vendor, target, action, param, enabled
Disabled rows are not judged (the engine ignores them).
"""
from __future__ import annotations

import re
from collections.abc import Mapping, Sequence

import pandas as pd

from core.config_loader import normalize_config
from core.dictionary import _parse_bool
from core.engine import IMPLEMENTED_ACTIONS
from store.rules_repo import rule_csvs_to_raw_sheets
from store.validators import Issue

DEFAULT_KEYWORD = "DEFAULT"
ALL_VENDORS = "ALL"
QTY_PLACEHOLDER = "{qty}"
REQUIRES_PARAM = frozenset({
    "REMOVE_TEXT", "REMOVE_REGEX", "REPLACE_REGEX_SUB", "MASK_TEXT", "CALC_UNIT",
    "GROUP_MULTIPLY", "APPEND_SUFFIX", "PREPEND_TEXT",
})
REGEX_ACTIONS = frozenset({"REMOVE_REGEX", "REPLACE_REGEX_SUB"})
RULE_COLUMN_LABELS: dict[str, str] = {
    "priority": "우선순위", "keyword": "키워드", "vendor": "양식명칭", "order": "순서",
    "target": "적용대상", "action": "ActionType", "param": "Parameter",
}


def _cell(values: Mapping[str, object], name: str) -> str:
    value = values.get(name)
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()


def is_int_text(text: str) -> bool:
    """True for '3', '-1' and '3.0' (an integer written as a float), False for '1.5' or ''."""
    try:
        return float(text).is_integer()
    except ValueError:
        return False


def int_value(text: str) -> int | None:
    return int(float(text)) if is_int_text(text) else None


def split_keywords(raw: str) -> list[str]:
    return [k.strip() for k in raw.split(",") if k.strip()]


def _active_rows(frame: pd.DataFrame) -> list[tuple[int, dict[str, object]]]:
    rows: list[tuple[int, dict[str, object]]] = []
    for position, (_, series) in enumerate(frame.iterrows()):
        values: dict[str, object] = {str(k): v for k, v in series.items()}
        if "enabled" in values and not _parse_bool(values["enabled"]):
            continue
        rows.append((position, values))
    return rows


# ---------------------------------------------------------------- ProductRoute

def _route_row_issues(position: int, values: Mapping[str, object], vendor_ids: Sequence[str]) -> list[Issue]:
    issues: list[Issue] = []
    if not is_int_text(_cell(values, "priority")):
        issues.append(Issue("error", position, "priority", "우선순위는 정수(예: 1, 2, 3)여야 합니다."))
    if not split_keywords(_cell(values, "keyword")):
        issues.append(Issue("error", position, "keyword", "키워드가 비어 있습니다. (기본값이면 DEFAULT)"))
    vendor = _cell(values, "vendor")
    if not vendor:
        issues.append(Issue("error", position, "vendor", "양식명칭이 비어 있습니다."))
    elif vendor not in vendor_ids:
        issues.append(Issue("error", position, "vendor", f"양식명칭 '{vendor}'을(를) 찾을 수 없습니다."))
    return issues


def _is_default(keyword: str) -> bool:
    return any(k.upper() == DEFAULT_KEYWORD for k in split_keywords(keyword))


def _default_issues(rows: list[tuple[int, dict[str, object]]]) -> list[Issue]:
    count = sum(1 for _, v in rows if _is_default(_cell(v, "keyword")))
    if count == 1:
        return []
    detail = "없습니다" if count == 0 else f"{count}개 있습니다"
    return [Issue(
        "error", None, "keyword",
        f"기본값(DEFAULT) 행이 정확히 1개 있어야 합니다. 지금은 {detail}. 기본값 행은 삭제하거나 끌 수 없습니다.",
    )]


def _ordered_keywords(rows: list[tuple[int, dict[str, object]]]) -> list[tuple[int, int, str]]:
    """(position, priority, keyword) of every non-DEFAULT keyword, in evaluation order."""
    entries: list[tuple[int, int, int, str]] = []
    for position, values in rows:
        priority = int_value(_cell(values, "priority"))
        if priority is None:
            continue
        for keyword in split_keywords(_cell(values, "keyword")):
            if keyword.upper() != DEFAULT_KEYWORD:
                entries.append((priority, position, len(entries), keyword))
    entries.sort()
    return [(position, priority, keyword) for priority, position, _i, keyword in entries]


def _keyword_warnings(rows: list[tuple[int, dict[str, object]]]) -> list[Issue]:
    ordered = _ordered_keywords(rows)
    issues: list[Issue] = []
    seen_pairs: set[tuple[str, str]] = set()
    for j, (pos_b, prio_b, kw_b) in enumerate(ordered):
        for pos_a, prio_a, kw_a in ordered[:j]:
            if pos_a == pos_b:
                continue
            if kw_a == kw_b:
                message = f"키워드 '{kw_a}'이(가) 여러 행(우선순위 {prio_a}, {prio_b})에 중복되어 있습니다."
            elif kw_a in kw_b:
                message = (
                    f"'{kw_a}'(우선순위 {prio_a})이(가) 먼저 검사되어 '{kw_b}'(우선순위 {prio_b})는 "
                    f"항상 '{kw_a}' 쪽으로 분류됩니다. '{kw_b}'는 절대 적용되지 않습니다 (가려짐)."
                )
            else:
                continue
            if (kw_a, kw_b) in seen_pairs:
                continue
            seen_pairs.add((kw_a, kw_b))
            issues += [Issue("warning", pos_b, "keyword", message), Issue("warning", pos_a, "keyword", message)]
    return issues


def validate_product_route(frame: pd.DataFrame, vendor_ids: Sequence[str]) -> list[Issue]:
    """Errors: priority, vendor, empty keyword, DEFAULT count. Warnings: duplicate / shadowed keywords."""
    rows = _active_rows(frame)
    issues: list[Issue] = []
    for position, values in rows:
        issues += _route_row_issues(position, values, vendor_ids)
    return issues + _default_issues(rows) + _keyword_warnings(rows)


# ---------------------------------------------------------------- OptionRules

def _regex_part(action: str, param: str) -> str:
    """The part of a regex-action Parameter that is compiled (REPLACE: before '///' or '||')."""
    if action != "REPLACE_REGEX_SUB":
        return param
    if "///" in param:
        return param.split("///", 1)[0].strip()
    if "||" in param:
        return param.split("||", 1)[0].strip()
    return param


def _action_issues(position: int, action: str, param: str) -> list[Issue]:
    issues: list[Issue] = []
    if action in REQUIRES_PARAM and not param:
        issues.append(Issue("error", position, "param", f"{action}은(는) Parameter가 비어 있으면 안 됩니다."))
    if action in REGEX_ACTIONS and param:
        pattern = _regex_part(action, param)
        try:
            re.compile(pattern)
        except re.error as e:
            issues.append(Issue("error", position, "param", f"정규식 패턴이 올바르지 않습니다: {e} (패턴: {pattern})"))
    if action == "REPLACE_REGEX_SUB" and param and "///" not in param and "||" not in param:
        issues.append(Issue(
            "warning", position, "param",
            "'패턴 /// 바꿀 말' 형식이 아니어서, 패턴에 맞는 부분이 그냥 지워집니다.",
        ))
    if action == "FORMAT_QTY" and param and QTY_PLACEHOLDER not in param:
        issues.append(Issue(
            "warning", position, "param",
            "{qty}가 없어서 서식이 무시되고 ' x수량개'로 붙습니다.",
        ))
    return issues


def _option_row_issues(position: int, values: Mapping[str, object], vendor_ids: Sequence[str]) -> list[Issue]:
    action = _cell(values, "action").upper()
    vendor = _cell(values, "vendor")
    target = _cell(values, "target")
    issues: list[Issue] = []
    if not is_int_text(_cell(values, "order")):
        issues.append(Issue("error", position, "order", "순서는 정수여야 합니다."))
    if action not in IMPLEMENTED_ACTIONS:
        issues.append(Issue("error", position, "action", f"ActionType '{action or '(비어 있음)'}'은(는) 지원하지 않습니다."))
    known = {v.strip().upper() for v in vendor_ids}
    if vendor.upper() != ALL_VENDORS and vendor.upper() not in known:
        issues.append(Issue("error", position, "vendor", f"양식명칭 '{vendor}'을(를) 찾을 수 없습니다. (ALL 또는 발주양식)"))
    if not target:
        issues.append(Issue("warning", position, "target", "적용대상이 비어 있습니다. 모든 상품에 적용하려면 ALL이라고 쓰세요."))
    elif "," in target:
        issues.append(Issue("warning", position, "target", "적용대상에는 문구를 하나만 쓸 수 있습니다. 쉼표로 여러 개를 쓰면 통째로 하나의 문구로 찾습니다."))
    return issues + _action_issues(position, action, _cell(values, "param"))


def validate_option_rules(frame: pd.DataFrame, vendor_ids: Sequence[str]) -> list[Issue]:
    """Validate the enabled rows of a canonical option-rules frame."""
    issues: list[Issue] = []
    for position, values in _active_rows(frame):
        issues += _option_row_issues(position, values, vendor_ids)
    return issues


# ---------------------------------------------------------------- final guard

def build_config(product_route_csv: str, option_rules_csv: str, layout: pd.DataFrame) -> dict:
    """The config dict the pipeline would use for these CSV texts (raises on unusable rules)."""
    raw = rule_csvs_to_raw_sheets(product_route_csv, option_rules_csv)
    return normalize_config({**raw, "OutputLayout": layout.copy()})


def final_guard(product_route_csv: str, option_rules_csv: str, layout: pd.DataFrame) -> list[Issue]:
    """Error when the new CSVs cannot be turned into a working config (always judged)."""
    try:
        build_config(product_route_csv, option_rules_csv, layout)
    except (ValueError, KeyError, TypeError, AttributeError, pd.errors.ParserError) as e:
        return [Issue("error", None, "", f"이 규칙으로는 주문을 처리할 수 없습니다. ({type(e).__name__}: {str(e)[:200]})")]
    return []
