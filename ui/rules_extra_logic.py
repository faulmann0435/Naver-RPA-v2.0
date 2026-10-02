"""Pure helpers of the 고급 설정 page besides the grids: change impact, rule test, settings.json."""
from __future__ import annotations

import json
import re
from collections import Counter
from dataclasses import dataclass
from itertools import zip_longest

import pandas as pd

from core.dictionary import DictionarySettings, ItemDictionary
from core.pipeline import process_orders
from ui.preview_logic import preview_bundle

UNCLASSIFIED_LABEL = "미분류"
IMPACT_COLUMNS = ["이전 양식", "이전 품목", "이후 양식", "이후 품목"]
SEPARATOR_CHOICES: tuple[str, ...] = (", ", " / ", " + ")
SETTING_LABELS = {"item_separator": "품목 구분 기호", "ignored_option_groups": "비교 제외 옵션 항목"}


# ---------------------------------------------------------------- change impact

@dataclass(frozen=True)
class ImpactResult:
    total_lines: int
    changed_lines: int
    table: pd.DataFrame  # IMPACT_COLUMNS; vendor and item text only (no customer data)


def output_lines(merged: pd.DataFrame | None) -> list[tuple[str, str]]:
    """(vendor, item text) of every output line."""
    if merged is None or merged.empty or "_VendorID" not in merged.columns:
        return []
    return [
        (UNCLASSIFIED_LABEL if str(v) == "Unclassified" else str(v), str(t))
        for v, t in zip(merged["_VendorID"], merged["processed_option"])
    ]


def _not_in(lines: list[tuple[str, str]], other: list[tuple[str, str]]) -> list[tuple[str, str]]:
    remaining = Counter(other)
    left: list[tuple[str, str]] = []
    for line in lines:
        if remaining[line] > 0:
            remaining[line] -= 1
        else:
            left.append(line)
    return left


def compare_outputs(before: pd.DataFrame | None, after: pd.DataFrame | None) -> ImpactResult:
    """Lines that exist only before / only after (matched as a multiset), paired in order."""
    old, new = output_lines(before), output_lines(after)
    removed, added = _not_in(old, new), _not_in(new, old)
    empty = ("", "")
    rows = [[*a, *b] for a, b in zip_longest(removed, added, fillvalue=empty)]
    return ImpactResult(len(old), max(len(removed), len(added)), pd.DataFrame(rows, columns=IMPACT_COLUMNS))


def run_impact(
    orders: pd.DataFrame, saved_config: dict, edited_config: dict,
    dictionary: ItemDictionary, settings: DictionarySettings,
) -> ImpactResult:
    """Process the same order file with the saved and the edited rules (same dictionary state)."""
    before = process_orders(orders.copy(), saved_config, dictionary, settings).merged
    after = process_orders(orders.copy(), edited_config, dictionary, settings).merged
    return compare_outputs(before, after)


# ---------------------------------------------------------------- rule test

@dataclass(frozen=True)
class LogStep:
    number: int
    action: str
    param: str
    applied: bool
    note: str  # Korean explanation of what happened
    vendor_rule: str
    target: str
    description: str


@dataclass(frozen=True)
class RuleTestResult:
    vendor: str
    text: str
    steps: list[LogStep]


_LOG_LINE = re.compile(r"^Row \d+ Rule #(\d+) \(Action: ([^,)]*)(?:, Param: (.*?))?\) -> (.*)$")
_SKIP_NOTES = {
    "NO (ApplyTo)": "건너뜀: 양식명칭이 다릅니다",
    "NO (TargetKeyword)": "건너뜀: 적용대상 문구가 상품명/옵션에 없습니다",
    "SKIP (lock)": "건너뜀: 무게 변환은 한 행에 한 번만 적용됩니다",
    "SKIP (qty_display_lock)": "건너뜀: 수량 표시는 한 행에 한 번만 적용됩니다",
}


def _note(tail: str) -> tuple[bool, str]:
    for key, note in _SKIP_NOTES.items():
        if key in tail:
            return False, note
    if "YES" in tail:
        result = tail.split("Result:", 1)[-1].strip()
        return True, f"적용 → {result}"
    return False, tail


def parse_rule_log(lines: list[str], option_rules: pd.DataFrame) -> list[LogStep]:
    """Turn the engine's debug lines into Korean steps, with the rule's own columns attached."""
    steps: list[LogStep] = []
    for line in lines:
        match = _LOG_LINE.match(line)
        if match is None:
            continue
        number = int(match.group(1))
        applied, note = _note(match.group(4))
        rule = option_rules.iloc[number - 1] if 0 < number <= len(option_rules) else pd.Series(dtype=object)
        steps.append(LogStep(
            number=number, action=match.group(2), param=(match.group(3) or "").strip("'\""), applied=applied,
            note=note, vendor_rule=str(rule.get("ApplyTo", "")), target=str(rule.get("TargetKeyword", "")),
            description=str(rule.get("Description", "") if pd.notna(rule.get("Description", "")) else ""),
        ))
    return steps


def run_rule_test(name: str, option: str, qty: int, config: dict) -> RuleTestResult | None:
    """One item through routing and the auto rules only (empty dictionary); None when nothing came out."""
    rows = pd.DataFrame({"상품번호": [""], "상품명": [name], "옵션정보": [option], "수량": [max(int(qty), 1)]})
    result = preview_bundle(rows, config, ItemDictionary.empty(), DictionarySettings())
    if not result.vendors:
        return None
    first = result.vendors[0]
    vendor = UNCLASSIFIED_LABEL if first.vendor == "Unclassified" else first.vendor
    return RuleTestResult(vendor, first.text, parse_rule_log(result.debug.get(0, []), config["OptionRules"]))


# ---------------------------------------------------------------- settings.json

def settings_values(data: dict) -> tuple[str, list[str]]:
    """(item separator, ignored option groups) of a parsed settings.json, with the defaults."""
    defaults = DictionarySettings()
    separator = str(data.get("item_separator", defaults.item_separator))
    groups = data.get("ignored_option_groups", list(defaults.ignored_option_groups))
    return separator, [str(g) for g in groups] if isinstance(groups, list) else []


def parse_settings(text: str | None) -> dict:
    """settings.json as a dict (missing / empty / not an object -> {})."""
    if text is None or not text.strip():
        return {}
    data = json.loads(text)
    return data if isinstance(data, dict) else {}


def clean_groups(groups: list[object]) -> list[str]:
    """Trimmed, non-empty, no duplicates, original order."""
    cleaned = (str(g).strip() for g in groups if g is not None and not (isinstance(g, float) and pd.isna(g)))
    return list(dict.fromkeys(g for g in cleaned if g))


def validate_settings(separator: str) -> list[str]:
    return [] if separator.strip() else ["품목 구분 기호가 비어 있습니다."]


def build_settings_text(existing: dict, separator: str, groups: list[str]) -> str:
    """settings.json text (same layout as the migration tool); other keys are kept."""
    data = {**existing, "ignored_option_groups": groups, "item_separator": separator}
    return json.dumps(data, ensure_ascii=False, indent=2) + "\n"
