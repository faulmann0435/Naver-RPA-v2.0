"""Pure helpers of the preview page (no streamlit): run one bundle through the real pipeline."""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from core.dictionary import DictionarySettings, ItemDictionary
from core.engine import apply_option_rules
from core.pipeline import process_orders
from core.router import route_vendor

PATH_DICT = "사전"
PATH_AUTO = "자동 규칙"
PREVIEW_COLUMNS = ["상품번호", "상품명", "옵션정보", "수량"]


@dataclass(frozen=True)
class VendorPreview:
    vendor: str
    text: str


@dataclass(frozen=True)
class PreviewResult:
    vendors: list[VendorPreview]
    paths: list[str]  # per input row: 사전 / 자동 규칙
    debug: dict[int, list[str]] = field(default_factory=dict)  # input row -> rule log (auto rows only)


def clean_rows(rows_df: pd.DataFrame) -> pd.DataFrame:
    """Rows with a product name or option; quantity coerced to an int >= 1."""
    frame = rows_df.reindex(columns=PREVIEW_COLUMNS).copy()
    for col in ("상품번호", "상품명", "옵션정보"):
        frame[col] = frame[col].map(lambda v: "" if v is None or pd.isna(v) else str(v).strip())
    qty = pd.to_numeric(frame["수량"], errors="coerce").fillna(1).astype(int)
    frame["수량"] = qty.clip(lower=1)
    return frame[(frame["상품명"] != "") | (frame["옵션정보"] != "")].reset_index(drop=True)


def _debug_log(row: pd.Series, config: dict) -> list[str]:
    one = route_vendor(pd.DataFrame([row]), config["ProductRoute"]).iloc[0]
    log: list[str] = []
    apply_option_rules(one, config["OptionRules"], row_index=0, debug_log=log)
    return log


def preview_bundle(
    rows_df: pd.DataFrame, config: dict, dictionary: ItemDictionary, settings: DictionarySettings
) -> PreviewResult:
    """Treat all rows as ONE shipping bundle and return the final 품목 text per vendor."""
    rows = clean_rows(rows_df)
    if rows.empty:
        return PreviewResult([], [])
    order = rows.copy()
    order["배송비 묶음번호"] = 1
    order["결제일"] = ""
    result = process_orders(order, config, dictionary, settings)
    merged = result.merged if result.merged is not None else pd.DataFrame()
    vendors = [
        VendorPreview(str(r["_VendorID"]), str(r["processed_option"])) for _, r in merged.iterrows()
    ]
    paths: list[str] = []
    debug: dict[int, list[str]] = {}
    for i, row in rows.iterrows():
        hit = dictionary.lookup(row["상품번호"], row["옵션정보"], settings)
        paths.append(PATH_DICT if hit is not None else PATH_AUTO)
        if hit is None:
            debug[int(str(i))] = _debug_log(row, config)
    return PreviewResult(vendors, paths, debug)
