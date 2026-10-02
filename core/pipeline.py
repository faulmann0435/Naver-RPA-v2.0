"""Full pipeline: filter -> route -> rules -> [dictionary override] -> sort -> merge -> sort -> export."""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from core.actions import _safe_int
from core.dictionary import DictionaryEntry, DictionarySettings, ItemDictionary
from core.engine import run_option_engine
from core.exporter import export_individual_files
from core.merger import (
    filter_instruction_rows,
    merge_orders,
    merge_orders_with_dictionary,
    sort_by_payment_date,
)
from core.option_key import make_option_key, normalize_product_no
from core.router import route_vendor

DICT_ENTRY_COL = "_dict_entry"
UNMATCHED_COLUMNS = [
    "product_no", "option_key", "상품명", "옵션정보", "qty_example",
    "vendor_suggestion", "display_suggestion", "count",
]
STAT_KEYS = ("rows_total", "rows_dictionary", "rows_auto", "rows_unclassified", "rows_needs_review")


@dataclass(frozen=True)
class ProcessResult:
    files: list[dict]
    unmatched: pd.DataFrame
    stats: dict[str, int]


def _column_values(df: pd.DataFrame, name: str) -> list:
    return df[name].tolist() if name in df.columns else [None] * len(df)


def _lookup_entries(
    df: pd.DataFrame, dictionary: ItemDictionary, settings: DictionarySettings
) -> list[DictionaryEntry | None]:
    if not len(dictionary):
        return [None] * len(df)
    return [
        dictionary.lookup(product_no, option, settings)
        for product_no, option in zip(_column_values(df, "상품번호"), _column_values(df, "옵션정보"))
    ]


def _unmatched_table(
    df: pd.DataFrame, entries: list[DictionaryEntry | None], settings: DictionarySettings
) -> pd.DataFrame:
    groups: dict[tuple[str, str], dict] = {}
    columns = {
        name: _column_values(df, name)
        for name in ("상품번호", "상품명", "옵션정보", "수량", "_VendorID", "processed_option")
    }
    for i, entry in enumerate(entries):
        product_no = normalize_product_no(columns["상품번호"][i])
        if entry is not None or not product_no:
            continue
        option_key = make_option_key(columns["옵션정보"][i], settings.ignored_option_groups)
        key = (product_no, option_key)
        if key in groups:
            groups[key]["count"] += 1
            continue
        option_raw = columns["옵션정보"][i]
        groups[key] = {
            "product_no": product_no,
            "option_key": option_key,
            "상품명": "" if columns["상품명"][i] is None else columns["상품명"][i],
            "옵션정보": "" if pd.isna(option_raw) else option_raw,
            "qty_example": _safe_int(columns["수량"][i], 1),
            "vendor_suggestion": columns["_VendorID"][i],
            "display_suggestion": columns["processed_option"][i],
            "count": 1,
        }
    return pd.DataFrame(list(groups.values()), columns=UNMATCHED_COLUMNS)


def _stats(df: pd.DataFrame, entries: list[DictionaryEntry | None]) -> dict[str, int]:
    matched = [e for e in entries if e is not None]
    return {
        "rows_total": len(df),
        "rows_dictionary": len(matched),
        "rows_auto": len(df) - len(matched),
        "rows_unclassified": int((df["_VendorID"] == "Unclassified").sum()) if "_VendorID" in df.columns else 0,
        "rows_needs_review": sum(1 for e in matched if e.needs_review),
    }


def process_orders(
    df,
    config,
    dictionary: ItemDictionary | None = None,
    settings: DictionarySettings | None = None,
) -> ProcessResult:
    """Same steps as the classic pipeline, with dictionary rows overriding vendor and display text."""
    product_route = config["ProductRoute"]
    option_rules = config["OptionRules"]
    dictionary = dictionary if dictionary is not None else ItemDictionary.empty()
    settings = settings if settings is not None else DictionarySettings()

    # 1. 필터링 (Filter)
    df = filter_instruction_rows(df)
    if df.empty:
        return ProcessResult(files=[], unmatched=pd.DataFrame(columns=UNMATCHED_COLUMNS),
                             stats=dict.fromkeys(STAT_KEYS, 0))
    if "구매자명" not in df.columns:
        df = df.copy()
        df["구매자명"] = ""

    # 2. 업체 분류 (Routing)
    df = route_vendor(df, product_route)

    # 3. 룰 적용 (Option Engine)
    df = run_option_engine(df, option_rules, debug_log=None)

    # 3b. 품목 사전 조회: 일치한 행은 업체를 덮어쓰고 표시 문구를 사전에서 가져온다.
    entries = _lookup_entries(df, dictionary, settings)
    unmatched = _unmatched_table(df, entries, settings)
    has_matches = any(e is not None for e in entries)
    if has_matches:
        df[DICT_ENTRY_COL] = pd.Series(entries, index=df.index, dtype=object)
        df["_VendorID"] = [e.vendor_id if e is not None else v for e, v in zip(entries, df["_VendorID"])]
    stats = _stats(df, entries)

    # 병합 전 정렬 (Pre-Merge Sort): 상품주문번호 오름차순
    sort_col = "상품주문번호"
    if sort_col in df.columns:
        df[sort_col] = df[sort_col].astype(str)
        df = df.sort_values(by=sort_col, ascending=True)

    # 4. 병합 (Merge)
    if has_matches:
        merged = merge_orders_with_dictionary(df, separator=settings.item_separator)
    else:
        merged = merge_orders(df, option_rules=option_rules)
    merged = merged.drop(columns=[DICT_ENTRY_COL], errors="ignore")

    # 5. 결제일 기준 최종 정렬 (Sort by Payment Date)
    if "결제일" in merged.columns:
        merged = sort_by_payment_date(merged)

    # 6. 파일 생성 (Export)
    return ProcessResult(files=export_individual_files(merged, config), unmatched=unmatched, stats=stats)


def process_all_data(df, config):
    """Classic entry point (no dictionary): filter -> route -> rules -> sort -> merge -> sort -> export."""
    return process_orders(df, config).files
