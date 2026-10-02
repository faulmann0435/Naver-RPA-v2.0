"""Pure helpers of the option form (no streamlit): form values, live results, inline validation."""
from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass

import pandas as pd

from core.dictionary import (
    DictionaryEntry,
    DictionarySettings,
    _parse_bool,
    _parse_float,
    render_entry,
)
from core.merger import _format_weight
from core.template import TemplateError, format_number
from store.validators import Issue, validate_new_entry

RESULT_QUANTITIES = (1, 2, 3)
APPEND_PREFIX = "묶음 끝에 붙음: "
RENDER_ERROR = "(표기 오류)"
_PLACEHOLDER_PRODUCT_NO = "0"


@dataclass(frozen=True)
class FormValues:
    template: str = ""
    qty1_template: str = ""
    vendor: str = ""
    sum_group: str = ""
    unit_weight_kg: float | None = None
    append_to_end: bool = False
    needs_review: bool = False
    enabled: bool = True


def _s(value: object) -> str:
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()


def values_from_row(row: Mapping[str, object]) -> FormValues:
    """Form values of one dictionary.csv row (text columns)."""
    return FormValues(
        template=_s(row.get("display_template")),
        qty1_template=_s(row.get("display_template_qty1")),
        vendor=_s(row.get("vendor_id")),
        sum_group=_s(row.get("sum_group")),
        unit_weight_kg=_parse_float(row.get("unit_weight_kg")),
        append_to_end=_parse_bool(row.get("append_to_end")),
        needs_review=_parse_bool(row.get("needs_review")),
        enabled=_parse_bool(row.get("enabled")),
    )


def weight_text(weight: float | None) -> str:
    return "" if weight is None else format_number(weight)


def row_fields(values: FormValues) -> dict[str, str]:
    """dictionary.csv columns (as text) that the form controls."""
    return {
        "vendor_id": values.vendor,
        "display_template": values.template.strip(),
        "display_template_qty1": values.qty1_template.strip(),
        "sum_group": values.sum_group.strip(),
        "unit_weight_kg": weight_text(values.unit_weight_kg),
        "append_to_end": "1" if values.append_to_end else "0",
        "needs_review": "1" if values.needs_review else "0",
        "enabled": "1" if values.enabled else "0",
    }


def render_values(values: FormValues, qty: int) -> str:
    """The purchase-order text for `qty` units (sum groups show the total weight)."""
    group = values.sum_group.strip()
    if group and values.unit_weight_kg:
        return f"{group} {_format_weight(values.unit_weight_kg * qty)}"
    entry = DictionaryEntry(
        product_no=_PLACEHOLDER_PRODUCT_NO, option_key="", vendor_id=values.vendor,
        display_template=values.template.strip(), display_template_qty1=values.qty1_template.strip(),
    )
    try:
        text = render_entry(entry, qty)
    except (TemplateError, ValueError):
        return RENDER_ERROR
    if not values.template.strip():
        return ""
    return APPEND_PREFIX + text if values.append_to_end else text


def result_lines(values: FormValues) -> list[tuple[int, str]]:
    return [(qty, render_values(values, qty)) for qty in RESULT_QUANTITIES]


def form_issues(
    values: FormValues, vendor_ids: Sequence[str], qty_example: int = 0, settings: DictionarySettings | None = None
) -> list[Issue]:
    """Validation of one entry (errors and warnings); the product number is not known here."""
    row = {"product_no": _PLACEHOLDER_PRODUCT_NO, **row_fields(values)}
    return validate_new_entry(row, qty_example, vendor_ids, settings or DictionarySettings())
