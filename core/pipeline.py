"""Full pipeline: filter -> route -> rules -> [dictionary override] -> sort -> merge -> sort -> export."""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from core.actions import _safe_int
from core.dictionary import DictionaryEntry, DictionarySettings, ItemDictionary
from core.engine import run_option_engine
from core.exporter import export_individual_files
from core.merger import (
    RECEIPT_DATE_COL,
    filter_instruction_rows,
    merge_orders,
    merge_orders_with_dictionary,
    sort_by_payment_date,
)
from core.option_key import make_option_key, normalize_product_no, receipt_date
from core.router import route_vendor

DICT_ENTRY_COL = "_dict_entry"
KIND_COL = "상품종류"
ADDON_KIND = "추가구성상품"
BUNDLE_COL = "배송비 묶음번호"
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
    merged: pd.DataFrame | None = None  # merged frame before export (used by the preview page)


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


def follow_main_vendor(df: pd.DataFrame) -> pd.DataFrame:
    """추가구성상품 rows take the form of their main product (same 묶음번호 + 상품번호).

    When the main product's rows went to several forms, an add-on already on one of them stays;
    otherwise the earliest 상품주문번호 wins. An add-on with no main row in its bundle keeps its own form,
    and so does one whose dictionary entry has own_vendor (a separately shipped product sold as an add-on)."""
    needed = (KIND_COL, BUNDLE_COL, "상품번호", "_VendorID")
    if any(c not in df.columns for c in needed):
        return df
    is_addon = df[KIND_COL].astype(str).str.strip() == ADDON_KIND
    if not is_addon.any():
        return df
    own = [bool(getattr(e, "own_vendor", False)) for e in df[DICT_ENTRY_COL]] if DICT_ENTRY_COL in df.columns else [False] * len(df)
    keys = list(zip(df[BUNDLE_COL].astype(str), df["상품번호"].map(normalize_product_no)))
    order = df["상품주문번호"].astype(str).tolist() if "상품주문번호" in df.columns else [""] * len(df)
    mains: dict[tuple[str, str], list[tuple[str, object]]] = {}  # key -> [(상품주문번호, vendor)]
    for key, number, vendor, addon in zip(keys, order, df["_VendorID"], is_addon):
        if not addon:
            mains.setdefault(key, []).append((number, vendor))

    def followed(key: tuple[str, str], vendor: object) -> object:
        candidates = mains.get(key)
        if not candidates or vendor in {v for _, v in candidates}:
            return vendor
        return min(candidates, key=lambda c: c[0])[1]

    out = df.copy()
    out["_VendorID"] = [
        followed(key, vendor) if addon and not keep else vendor
        for key, vendor, addon, keep in zip(keys, df["_VendorID"], is_addon, own)
    ]
    return out


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
        # 예약발송 상품의 수령일은 사전 표기가 아니라 그 주문의 옵션에서 가져온다.
        df[RECEIPT_DATE_COL] = [receipt_date(o) for o in _column_values(df, "옵션정보")]

    # 3c. 추가구성상품(아이스팩·올리브오일·젓갈 등)은 본 상품의 발주양식을 따라간다.
    #     단, 사전에서 own_vendor로 표시한 항목(따로 파는 참가리비 등)은 자기 발주양식을 유지한다.
    df = follow_main_vendor(df)
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
    merged = merged.drop(columns=[DICT_ENTRY_COL, RECEIPT_DATE_COL], errors="ignore")

    # 5. 결제일 기준 최종 정렬 (Sort by Payment Date)
    if "결제일" in merged.columns:
        merged = sort_by_payment_date(merged)

    # 6. 파일 생성 (Export)
    return ProcessResult(
        files=export_individual_files(merged, config), unmatched=unmatched, stats=stats, merged=merged
    )


def process_all_data(df, config):
    """Classic entry point (no dictionary): filter -> route -> rules -> sort -> merge -> sort -> export."""
    return process_orders(df, config).files
