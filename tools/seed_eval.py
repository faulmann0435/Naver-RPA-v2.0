"""Compare pipeline output (without / with the seeded dictionary) to hand-written PO texts."""
from __future__ import annotations

import re
from dataclasses import dataclass
from io import BytesIO

import pandas as pd

from core.dictionary import DictionarySettings, ItemDictionary
from core.pipeline import process_orders
from tools.seed_data import PoBundle

NAME_SOURCE = "수취인명"
ITEM_SOURCE = "processed_option"
_SPACES = re.compile(r"\s+")


@dataclass(frozen=True)
class BundleResult:
    index: int
    hand: str
    a_texts: tuple[str, ...]
    b_texts: tuple[str, ...]
    exact_a: bool
    exact_b: bool
    norm_a: bool
    norm_b: bool


def normalize_item_text(text: str) -> str:
    return _SPACES.sub(" ", text.replace("★ ", "★")).strip()


def _header_for(layout: pd.DataFrame, vendor: str, source: str, default: str) -> str:
    rows = layout[layout["VendorID"].astype(str).str.strip() == vendor]
    hit = rows[rows["SourceCol"].astype(str).str.strip() == source]
    if hit.empty or pd.isna(hit["HeaderName"].iloc[0]):
        return default
    return str(hit["HeaderName"].iloc[0]).strip()


def item_texts_by_recipient(files: list[dict], layout: pd.DataFrame) -> dict[str, list[str]]:
    """recipient name -> item texts over all vendor output files (columns found through OutputLayout)."""
    texts: dict[str, list[str]] = {}
    for item in files:
        vendor = str(item["vendor"]).strip()
        name_col = _header_for(layout, vendor, NAME_SOURCE, NAME_SOURCE)
        item_col = _header_for(layout, vendor, ITEM_SOURCE, ITEM_SOURCE)
        frame = pd.read_excel(BytesIO(item["data"].getvalue()), dtype=object, keep_default_na=False)
        if name_col not in frame.columns or item_col not in frame.columns:
            continue
        for name, text in zip(frame[name_col], frame[item_col]):
            texts.setdefault(str(name).strip(), []).append(str(text))
    return texts


def _matches(texts: list[str], hand: str, normalize: bool) -> bool:
    if normalize:
        target = normalize_item_text(hand)
        return any(normalize_item_text(t) == target for t in texts)
    return any(t == hand for t in texts)


def evaluate_pair(
    orders: pd.DataFrame,
    bundles: list[PoBundle],
    config: dict,
    dictionary: ItemDictionary,
    settings: DictionarySettings,
) -> list[BundleResult]:
    """Run the pipeline without (A) and with (B) the dictionary and score every PO bundle."""
    layout = config["OutputLayout"]
    out_a = item_texts_by_recipient(process_orders(orders.copy(), config, ItemDictionary.empty(), settings).files, layout)
    out_b = item_texts_by_recipient(process_orders(orders.copy(), config, dictionary, settings).files, layout)
    results = []
    for bundle in bundles:
        a_texts = out_a.get(bundle.name, [])
        b_texts = out_b.get(bundle.name, [])
        results.append(BundleResult(
            index=bundle.index,
            hand=bundle.hand,
            a_texts=tuple(a_texts),
            b_texts=tuple(b_texts),
            exact_a=_matches(a_texts, bundle.hand, False),
            exact_b=_matches(b_texts, bundle.hand, False),
            norm_a=_matches(a_texts, bundle.hand, True),
            norm_b=_matches(b_texts, bundle.hand, True),
        ))
    return results
