"""Item dictionary (품목 사전): human-written supplier form + display text per (상품번호, option)."""
from __future__ import annotations

import json
import math
from dataclasses import dataclass
from typing import Any

import pandas as pd

from core.option_key import make_option_key, normalize_product_no
from core.template import has_qty_placeholder, render, strip_invisible

DEFAULT_IGNORED_GROUPS: tuple[str, ...] = ("수령일 선택 (도착시간 지정불가)",)
DEFAULT_SEPARATOR = " / "
DEFAULT_WORKERS: tuple[str, ...] = ("사장님", "사장님을노리는님", "개발자")
NAVER_CHANNEL = "naver"
_TRUE_TEXT = frozenset({"1", "true", "t", "y", "yes", "예", "o"})


@dataclass(frozen=True)
class DictionaryEntry:
    product_no: str
    option_key: str
    vendor_id: str
    display_template: str
    display_template_qty1: str = ""
    sum_group: str = ""
    unit_weight_kg: float | None = None
    append_to_end: bool = False
    needs_review: bool = False
    product_name_ref: str = ""
    option_raw_ref: str = ""

    @property
    def key(self) -> tuple[str, str]:
        return (self.product_no, self.option_key)


def parse_workers(value: object) -> tuple[str, ...]:
    """Trimmed, non-blank, no duplicates (order kept); the default list when nothing is left."""
    if not isinstance(value, list):
        return DEFAULT_WORKERS
    names = (str(v).strip() for v in value if v is not None)
    cleaned = tuple(dict.fromkeys(n for n in names if n))
    return cleaned or DEFAULT_WORKERS


@dataclass(frozen=True)
class DictionarySettings:
    ignored_option_groups: tuple[str, ...] = DEFAULT_IGNORED_GROUPS
    item_separator: str = DEFAULT_SEPARATOR
    workers: tuple[str, ...] = DEFAULT_WORKERS

    @classmethod
    def from_json_text(cls, text: str) -> DictionarySettings:
        data: Any = json.loads(text) if text.strip() else {}
        if not isinstance(data, dict):
            data = {}
        groups = data.get("ignored_option_groups", DEFAULT_IGNORED_GROUPS)
        separator = data.get("item_separator", DEFAULT_SEPARATOR)
        return cls(
            ignored_option_groups=tuple(str(g) for g in groups),
            item_separator=str(separator),
            workers=parse_workers(data.get("workers")),
        )


def _missing(value: object) -> bool:
    return value is None or (isinstance(value, float) and math.isnan(value))


def _text(value: object) -> str:
    return "" if _missing(value) else str(value).strip()


def _parse_bool(value: object) -> bool:
    if _missing(value):
        return False
    if isinstance(value, (bool, int, float)):
        return bool(value)
    return str(value).strip().lower() in _TRUE_TEXT


def _parse_float(value: object) -> float | None:
    text = _text(value)
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return None if math.isnan(number) else number


def _cell(row: pd.Series, name: str) -> object:
    return row[name] if name in row.index else None


def _entry_from_row(row: pd.Series, settings: DictionarySettings) -> DictionaryEntry | None:
    channel = _text(_cell(row, "channel"))
    if channel and channel != NAVER_CHANNEL:
        return None
    if "enabled" in row.index and not _parse_bool(row["enabled"]):
        return None
    vendor_id = _text(_cell(row, "vendor_id"))
    product_no = normalize_product_no(_cell(row, "product_no"))
    if not vendor_id or not product_no:
        return None
    return DictionaryEntry(
        product_no=product_no,
        option_key=make_option_key(_text(_cell(row, "option_key")), settings.ignored_option_groups),
        vendor_id=vendor_id,
        display_template=_text(_cell(row, "display_template")),
        display_template_qty1=_text(_cell(row, "display_template_qty1")),
        sum_group=_text(_cell(row, "sum_group")),
        unit_weight_kg=_parse_float(_cell(row, "unit_weight_kg")),
        append_to_end=_parse_bool(_cell(row, "append_to_end")),
        needs_review=_parse_bool(_cell(row, "needs_review")),
        product_name_ref=_text(_cell(row, "product_name_ref")),
        option_raw_ref=_text(_cell(row, "option_raw_ref")),
    )


class ItemDictionary:
    """Lookup table keyed by (normalized 상품번호, normalized option)."""

    def __init__(self, df: pd.DataFrame | None = None, settings: DictionarySettings | None = None) -> None:
        self._entries: dict[tuple[str, str], DictionaryEntry] = {}
        self.duplicates: list[tuple[str, str]] = []
        if df is None:
            return
        active = settings or DictionarySettings()
        for _, row in df.iterrows():
            entry = _entry_from_row(row, active)
            if entry is None:
                continue
            if entry.key in self._entries and entry.key not in self.duplicates:
                self.duplicates.append(entry.key)
            self._entries[entry.key] = entry

    @classmethod
    def empty(cls) -> ItemDictionary:
        return cls(None)

    def __len__(self) -> int:
        return len(self._entries)

    @property
    def entries(self) -> list[DictionaryEntry]:
        return list(self._entries.values())

    def lookup(
        self, product_no: object, option_text: object, settings: DictionarySettings
    ) -> DictionaryEntry | None:
        if not self._entries:
            return None
        key = (
            normalize_product_no(product_no),
            make_option_key(option_text, settings.ignored_option_groups),
        )
        return self._entries.get(key)


def render_entry(entry: DictionaryEntry, qty: int) -> str:
    """Display text for `qty` units; plain templates get a ' (xN)' suffix when qty > 1."""
    use_qty1 = qty == 1 and bool(entry.display_template_qty1)
    template = entry.display_template_qty1 if use_qty1 else entry.display_template
    text = strip_invisible(render(template, qty))
    if not has_qty_placeholder(template) and qty > 1:
        text = f"{text} (x{qty})"
    return text
