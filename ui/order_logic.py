"""Pure helpers of the order page (no streamlit): suggestions for unmatched rows and registration."""
from __future__ import annotations

import hashlib
from collections.abc import Sequence
from dataclasses import dataclass, field

import pandas as pd

from core.dictionary import (
    DictionaryEntry,
    DictionarySettings,
    _parse_float,
    render_entry,
)
from core.merger import _format_weight
from core.seed import probe_row, strip_ignored_groups
from store.base import Author, DataStore
from store.dictionary_repo import AppendResult, append_entries
from store.validators import Issue, validate_new_entry
from ui.option_logic import FormValues, form_issues, row_fields

UNCLASSIFIED = "Unclassified"
COL_CHECK = "등록"
COL_NAME = "상품명"
COL_OPTION = "옵션정보"
COL_QTY = "수량 예시"
COL_COUNT = "건수"
COL_VENDOR = "발주양식"
COL_TEMPLATE = "발주서 표기"
COL_QTY1 = "수량 1일 때 표기"
COL_GROUP = "합산 이름"
COL_WEIGHT = "1개당 무게(kg)"
SUGGESTION_COLUMNS = [
    COL_CHECK, COL_NAME, COL_OPTION, COL_QTY, COL_COUNT, COL_VENDOR,
    COL_TEMPLATE, COL_QTY1, COL_GROUP, COL_WEIGHT, "product_no", "option_key",
]
DISABLED_COLUMNS = [COL_NAME, COL_OPTION, COL_QTY, COL_COUNT]


def _text(value: object) -> str:
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value)


def _vendor_for(suggestion: object, vendor_ids: Sequence[str]) -> str:
    text = _text(suggestion).strip()
    if text in vendor_ids:
        return text
    return vendor_ids[0] if vendor_ids else text


def _suggestion_row(
    row: pd.Series, config: dict, vendor_ids: Sequence[str], settings: DictionarySettings
) -> dict:
    vendor = _vendor_for(row["vendor_suggestion"], vendor_ids)
    name, option = _text(row["상품명"]), _text(row["옵션정보"])
    probe = probe_row(
        name, strip_ignored_groups(option, settings.ignored_option_groups), vendor, config["OptionRules"]
    )
    template, group, weight = "", "", None
    if probe.sum_group:
        group, weight = probe.sum_group, probe.unit_weight_kg
    elif not probe.failed:
        template = probe.template
    else:
        template = _text(row["display_suggestion"]).strip()
    return {
        COL_CHECK: False, COL_NAME: name, COL_OPTION: option, COL_QTY: int(row["qty_example"]),
        COL_COUNT: int(row["count"]), COL_VENDOR: vendor, COL_TEMPLATE: template, COL_QTY1: "",
        COL_GROUP: group, COL_WEIGHT: weight, "product_no": row["product_no"], "option_key": row["option_key"],
    }


def build_suggestions(
    unmatched: pd.DataFrame,
    config: dict,
    vendor_ids: Sequence[str],
    settings: DictionarySettings | None = None,
) -> pd.DataFrame:
    """One editable suggestion row per unmatched (상품번호, 옵션) key."""
    active = settings or DictionarySettings()
    rows = [_suggestion_row(row, config, vendor_ids, active) for _, row in unmatched.iterrows()]
    frame = pd.DataFrame(rows, columns=SUGGESTION_COLUMNS)
    frame[COL_WEIGHT] = pd.to_numeric(frame[COL_WEIGHT], errors="coerce")
    return frame.astype({COL_CHECK: bool, COL_QTY: int, COL_COUNT: int})


def _weight_text(value: object) -> str:
    number = _parse_float(value)
    return "" if number is None else f"{number:g}"


def edited_to_row(edited_row: pd.Series) -> dict:
    """A suggestion row in dictionary.csv column terms."""
    return {
        "product_no": _text(edited_row["product_no"]),
        "option_key": _text(edited_row["option_key"]),
        "product_name_ref": _text(edited_row[COL_NAME]),
        "option_raw_ref": _text(edited_row[COL_OPTION]),
        "vendor_id": _text(edited_row[COL_VENDOR]),
        "display_template": _text(edited_row[COL_TEMPLATE]).strip(),
        "display_template_qty1": _text(edited_row[COL_QTY1]).strip(),
        "sum_group": _text(edited_row[COL_GROUP]).strip(),
        "unit_weight_kg": _weight_text(edited_row[COL_WEIGHT]),
    }


def checked_rows(edited: pd.DataFrame) -> pd.DataFrame:
    return edited[edited[COL_CHECK].fillna(False).astype(bool)]


def _entry_of(row: dict) -> DictionaryEntry:
    return DictionaryEntry(
        product_no=row["product_no"], option_key=row["option_key"], vendor_id=row["vendor_id"],
        display_template=row["display_template"], display_template_qty1=row["display_template_qty1"],
        sum_group=row["sum_group"], unit_weight_kg=_parse_float(row["unit_weight_kg"]),
    )


def render_example(row: dict, qty: int) -> str:
    """What the purchase order would show for `qty` units of this new entry."""
    entry = _entry_of(row)
    if entry.sum_group and entry.unit_weight_kg:
        return f"{entry.sum_group} {_format_weight(entry.unit_weight_kg * qty)}"
    return render_entry(entry, qty)


def row_issues(
    edited_row: pd.Series, vendor_ids: Sequence[str], settings: DictionarySettings
) -> list[Issue]:
    return validate_new_entry(edited_to_row(edited_row), int(edited_row[COL_QTY]), vendor_ids, settings)


def selection_preview(
    edited: pd.DataFrame, vendor_ids: Sequence[str], settings: DictionarySettings
) -> pd.DataFrame:
    """For checked rows: result at qty 1 and 2, or the validation error text."""
    records = []
    for _, row in checked_rows(edited).iterrows():
        issues = row_issues(row, vendor_ids, settings)
        errors = [i.message for i in issues if i.level == "error"]
        values = edited_to_row(row)
        records.append({
            COL_NAME: row[COL_NAME], COL_OPTION: row[COL_OPTION],
            "수량 1": "" if errors else render_example(values, 1),
            "수량 2": "" if errors else render_example(values, 2),
            "오류": " / ".join(errors),
        })
    return pd.DataFrame(records, columns=[COL_NAME, COL_OPTION, "수량 1", "수량 2", "오류"])


@dataclass(frozen=True)
class RegisterResult:
    status: str  # "saved" | "errors" | "needs_confirm" | "nothing_selected"
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    appended: AppendResult | None = None


def _label(row: pd.Series) -> str:
    return f"[{row[COL_NAME]} / {row[COL_OPTION] or '옵션 없음'}]"


def register_rows(
    store: DataStore,
    edited: pd.DataFrame,
    user: Author,
    settings: DictionarySettings,
    vendor_ids: Sequence[str],
    allow_warnings: bool = False,
) -> RegisterResult:
    """Validate the checked rows and append them to the dictionary. Errors block; warnings need confirmation."""
    selected = checked_rows(edited)
    if selected.empty:
        return RegisterResult("nothing_selected")
    errors: list[str] = []
    warnings: list[str] = []
    for _, row in selected.iterrows():
        for issue in row_issues(row, vendor_ids, settings):
            (errors if issue.level == "error" else warnings).append(f"{_label(row)} {issue.message}")
    if errors:
        return RegisterResult("errors", errors, warnings)
    if warnings and not allow_warnings:
        return RegisterResult("needs_confirm", [], warnings)
    rows = [edited_to_row(row) for _, row in selected.iterrows()]
    appended = append_entries(store, rows, user, settings)
    return RegisterResult("saved", [], warnings, appended)


# ---------------------------------------------------------------- per-item forms (option form)

@dataclass(frozen=True)
class UnmatchedItem:
    product_no: str
    option_key: str
    name: str
    option: str
    qty_example: int
    count: int


def unmatched_items(suggestions: pd.DataFrame) -> list[UnmatchedItem]:
    return [
        UnmatchedItem(
            _text(r["product_no"]), _text(r["option_key"]), _text(r[COL_NAME]), _text(r[COL_OPTION]),
            int(r[COL_QTY]), int(r[COL_COUNT]),
        )
        for _, r in suggestions.iterrows()
    ]


def suggestion_values(suggestions: pd.DataFrame, position: int) -> FormValues:
    """The prefilled form values of one suggestion row."""
    row = suggestions.iloc[position]
    return FormValues(
        template=_text(row[COL_TEMPLATE]).strip(), qty1_template=_text(row[COL_QTY1]).strip(),
        vendor=_text(row[COL_VENDOR]), sum_group=_text(row[COL_GROUP]).strip(),
        unit_weight_kg=_parse_float(row[COL_WEIGHT]), needs_review=True,
    )


def item_prefix(item: UnmatchedItem) -> str:
    digest = hashlib.sha1(item.option_key.encode("utf-8")).hexdigest()[:10]
    return f"um_{item.product_no}_{digest}"


def item_row(item: UnmatchedItem, values: FormValues) -> dict:
    """An item with its form values in dictionary.csv column terms."""
    fields = row_fields(values)
    return {
        "product_no": item.product_no, "option_key": item.option_key, "product_name_ref": item.name,
        "option_raw_ref": item.option, **{k: fields[k] for k in (
            "vendor_id", "display_template", "display_template_qty1", "sum_group", "unit_weight_kg",
            "append_to_end",
        )},
    }


def register_items(
    store: DataStore,
    selected: Sequence[tuple[UnmatchedItem, FormValues]],
    user: Author,
    settings: DictionarySettings,
    vendor_ids: Sequence[str],
    allow_warnings: bool = False,
) -> RegisterResult:
    """Like register_rows, for form values: errors block, warnings need confirmation."""
    if not selected:
        return RegisterResult("nothing_selected")
    errors: list[str] = []
    warnings: list[str] = []
    for item, values in selected:
        for issue in form_issues(values, vendor_ids, item.qty_example, settings):
            label = f"[{item.name} / {item.option or '옵션 없음'}] {issue.message}"
            (errors if issue.level == "error" else warnings).append(label)
    if errors:
        return RegisterResult("errors", errors, warnings)
    if warnings and not allow_warnings:
        return RegisterResult("needs_confirm", [], warnings)
    appended = append_entries(store, [item_row(i, v) for i, v in selected], user, settings)
    return RegisterResult("saved", [], warnings, appended)
