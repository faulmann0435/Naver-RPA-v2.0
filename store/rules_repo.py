"""Rule tables (ProductRoute / OptionRules) stored as CSV files in the data repository."""
from __future__ import annotations

import json
from dataclasses import dataclass

import pandas as pd

from core.config_loader import normalize_config, read_config_sheets
from core.engine import IMPLEMENTED_ACTIONS
from store.base import DataStore, StoreError
from store.csv_codec import from_csv_text, to_csv_text

PRODUCT_ROUTE_FILE = "product_route.csv"
OPTION_RULES_FILE = "option_rules.csv"
DICTIONARY_FILE = "dictionary.csv"
SETTINGS_FILE = "settings.json"

DEFAULT_CHANNEL = "naver"
META_COLUMNS = ["channel", "enabled", "updated_at", "updated_by"]
DICTIONARY_COLUMNS = [
    "channel", "product_no", "option_key", "product_name_ref", "option_raw_ref", "vendor_id",
    "display_template", "display_template_qty1", "sum_group", "unit_weight_kg", "append_to_end",
    "needs_review", "enabled", "source", "last_seen_at", "updated_at", "updated_by",
]
DEFAULT_SETTINGS = {
    "ignored_option_groups": ["수령일 선택 (도착시간 지정불가)"],
    "item_separator": " / ",
}


def _action_column(df: pd.DataFrame) -> str | None:
    """Find the ActionType column by header (same hints as config normalization)."""
    for col in df.columns:
        name = str(col).strip()
        if "ActionType" in name or "Action" in name or "명령" in name:
            return str(col)
    return None


def _drop_empty_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    keep = [c for c in df.columns if not (str(c).startswith("Unnamed") and df[c].isna().all())]
    return df[keep]


def _enabled_flags(df: pd.DataFrame, implemented_only: bool) -> list[int]:
    if not implemented_only:
        return [1] * len(df)
    col = _action_column(df)
    if col is None:
        return [1] * len(df)
    return [
        1 if str(v if pd.notna(v) else "").strip().upper() in IMPLEMENTED_ACTIONS else 0
        for v in df[col]
    ]


def _with_meta(df: pd.DataFrame, migrated_by: str, now: str, implemented_only: bool) -> pd.DataFrame:
    out = _drop_empty_unnamed(df).copy()
    out["channel"] = DEFAULT_CHANNEL
    out["enabled"] = _enabled_flags(out, implemented_only)
    out["updated_at"] = now
    out["updated_by"] = migrated_by
    return out


def sheets_to_rule_csvs(raw: dict[str, pd.DataFrame], migrated_by: str, now: str) -> dict[str, str]:
    """Convert raw config sheets into the files of the data repository."""
    route = _with_meta(raw["ProductRoute"], migrated_by, now, implemented_only=False)
    rules = _with_meta(raw["OptionRules"], migrated_by, now, implemented_only=True)
    return {
        PRODUCT_ROUTE_FILE: to_csv_text(route),
        OPTION_RULES_FILE: to_csv_text(rules),
        DICTIONARY_FILE: to_csv_text(pd.DataFrame(columns=DICTIONARY_COLUMNS)),
        SETTINGS_FILE: json.dumps(DEFAULT_SETTINGS, ensure_ascii=False, indent=2) + "\n",
    }


def _csv_to_raw(text: str) -> pd.DataFrame:
    df = from_csv_text(text)
    if "channel" in df.columns:
        df = df[df["channel"] == DEFAULT_CHANNEL]
    if "enabled" in df.columns:
        df = df[df["enabled"] == 1]
    return df.drop(columns=[c for c in META_COLUMNS if c in df.columns]).reset_index(drop=True)


def rule_csvs_to_raw_sheets(product_route_csv: str, option_rules_csv: str) -> dict[str, pd.DataFrame]:
    """Parse rule CSVs back into raw sheets (naver + enabled rows, META columns dropped)."""
    return {
        "ProductRoute": _csv_to_raw(product_route_csv),
        "OptionRules": _csv_to_raw(option_rules_csv),
    }


@dataclass(frozen=True)
class RulesSnapshot:
    """Raw rule sheets plus the shas they were read at."""

    product_route_raw: pd.DataFrame
    option_rules_raw: pd.DataFrame
    product_route_sha: str
    option_rules_sha: str


def _missing_message(name: str) -> str:
    return f"데이터 저장소에 규칙 파일 '{name}'이(가) 없습니다. 먼저 이관 도구를 실행하세요."


def load_rules(store: DataStore) -> RulesSnapshot:
    """Read both rule files from the store."""
    route = store.read_text(PRODUCT_ROUTE_FILE)
    rules = store.read_text(OPTION_RULES_FILE)
    if route is None:
        raise StoreError(_missing_message(PRODUCT_ROUTE_FILE))
    if rules is None:
        raise StoreError(_missing_message(OPTION_RULES_FILE))
    sheets = rule_csvs_to_raw_sheets(route.content, rules.content)
    return RulesSnapshot(
        product_route_raw=sheets["ProductRoute"],
        option_rules_raw=sheets["OptionRules"],
        product_route_sha=route.sha,
        option_rules_sha=rules.sha,
    )


def load_config_from_store(store: DataStore, config_path: str, password: str | None = None) -> dict:
    """Same result shape as load_config, with rules from the store and OutputLayout from config.xlsx."""
    layout = read_config_sheets(config_path, password)["OutputLayout"]
    rules = load_rules(store)
    return normalize_config(
        {
            "ProductRoute": rules.product_route_raw,
            "OptionRules": rules.option_rules_raw,
            "OutputLayout": layout,
        }
    )
