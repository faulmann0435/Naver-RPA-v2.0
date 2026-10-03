"""Build the draft dictionary rows from order data, option rules and past-PO evidence."""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd

from core.dictionary import DictionaryEntry, DictionarySettings
from core.option_key import make_option_key, normalize_product_no
from core.router import route_vendor
from core.seed import (
    HandEvidence,
    apply_evidence,
    probe_entry,
    probe_row,
    strip_ignored_groups,
    verify_probe,
)
from core.template import clean_display_text
from store.rules_repo import DICTIONARY_COLUMNS
from tools.seed_data import clean_text

UNCLASSIFIED = "Unclassified"
MUNGE_VENDOR = "킹선어, 대게선어, 깐멍게 발주양식"
MUNGE_KEYWORD = "멍게"
SOURCE_PAST_PO = "past_po"
SOURCE_AUTO = "auto_rule"
SKIP_NO_PRODUCT_NO = "상품번호 없음"
SKIP_UNCLASSIFIED = "Unclassified"
SKIP_MUNGE = "멍게(사용자 판단 대기)"

Key = tuple[str, str]


@dataclass(frozen=True)
class SeedItem:
    entry: DictionaryEntry
    source: str


@dataclass(frozen=True)
class KeyNote:
    """A key that needs attention in the report (product-level fields only)."""

    product_no: str
    product_name: str
    option_raw: str


@dataclass(frozen=True)
class SeedResult:
    items: list[SeedItem]
    skipped: dict[str, int]
    probe_failures: list[KeyNote] = field(default_factory=list)
    verify_failures: list[KeyNote] = field(default_factory=list)
    conflicts: list[KeyNote] = field(default_factory=list)
    rejected_evidence: int = 0
    q1_differs: int = 0


def distinct_keys(orders: pd.DataFrame, settings: DictionarySettings) -> dict[Key, tuple[str, str]]:
    """(product_no, option_key) -> (상품명, 옵션정보) of the first row seen."""
    firsts: dict[Key, tuple[str, str]] = {}
    for product_no, name, option in zip(orders["상품번호"], orders["상품명"], orders["옵션정보"]):
        key = (
            normalize_product_no(product_no),
            make_option_key(option, settings.ignored_option_groups),
        )
        firsts.setdefault(key, (clean_text(name), clean_text(option)))
    return firsts


def _skip_reason(key: Key, vendor: str, name: str, option: str) -> str | None:
    if not key[0]:
        return SKIP_NO_PRODUCT_NO
    if vendor == UNCLASSIFIED:
        return SKIP_UNCLASSIFIED
    if vendor == MUNGE_VENDOR and (MUNGE_KEYWORD in name or MUNGE_KEYWORD in option):
        return SKIP_MUNGE
    return None


def _route_all(firsts: dict[Key, tuple[str, str]], product_route: pd.DataFrame) -> list[str]:
    frame = pd.DataFrame(
        {"상품명": [v[0] for v in firsts.values()], "옵션정보": [v[1] for v in firsts.values()]}
    )
    return [str(v) for v in route_vendor(frame, product_route)["_VendorID"]]


def build_seed(
    orders: pd.DataFrame,
    config: dict,
    evidence: dict[Key, list[HandEvidence]],
    settings: DictionarySettings,
) -> SeedResult:
    firsts = distinct_keys(orders, settings)
    vendors = _route_all(firsts, config["ProductRoute"])
    items: list[SeedItem] = []
    skipped: Counter[str] = Counter()
    probe_failures: list[KeyNote] = []
    verify_failures: list[KeyNote] = []
    conflicts: list[KeyNote] = []
    rejected = 0
    q1_differs = 0
    for (key, (name, option)), vendor in zip(firsts.items(), vendors):
        reason = _skip_reason(key, vendor, name, option)
        if reason:
            skipped[reason] += 1
            continue
        note = KeyNote(key[0], name, option)
        probe = probe_row(name, strip_ignored_groups(option, settings.ignored_option_groups), vendor, config["OptionRules"])
        base = probe_entry(probe, key[0], key[1], vendor, name, option)
        if probe.failed:
            probe_failures.append(note)
        q1_differs += int(probe.q1_differs)
        if not verify_probe(base, probe):
            verify_failures.append(note)
        result = apply_evidence(base, evidence.get(key, []))
        rejected += result.rejected
        if result.conflict:
            conflicts.append(note)
        items.append(SeedItem(result.entry, SOURCE_PAST_PO if result.used else SOURCE_AUTO))
    return SeedResult(items, dict(skipped), probe_failures, verify_failures, conflicts, rejected, q1_differs)


def _row(item: SeedItem, now: str) -> dict[str, object]:
    entry = item.entry
    return {
        "channel": "naver",
        "product_no": entry.product_no,
        "option_key": entry.option_key,
        "product_name_ref": entry.product_name_ref,
        "option_raw_ref": entry.option_raw_ref,
        "vendor_id": entry.vendor_id,
        "display_template": clean_display_text(entry.display_template),
        "display_template_qty1": "",
        "sum_group": clean_display_text(entry.sum_group),
        "unit_weight_kg": "" if entry.unit_weight_kg is None else entry.unit_weight_kg,
        "append_to_end": 0,
        "needs_review": 1,
        "enabled": 1,
        "source": item.source,
        "last_seen_at": "",
        "updated_at": now,
        "updated_by": "seed",
        "created_at": now,
    }


def dictionary_frame(items: list[SeedItem], now: datetime) -> pd.DataFrame:
    """Dictionary table with exactly DICTIONARY_COLUMNS, sorted by vendor, product name, raw option."""
    stamp = now.isoformat(timespec="seconds")
    rows = [_row(item, stamp) for item in items]
    frame = pd.DataFrame(rows, columns=DICTIONARY_COLUMNS)
    return frame.sort_values(
        ["vendor_id", "product_name_ref", "option_raw_ref"], kind="stable"
    ).reset_index(drop=True)
