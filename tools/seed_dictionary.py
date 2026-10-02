"""Build a DRAFT item dictionary from historical orders and score it against past hand-written POs.

Usage:
    python -m tools.seed_dictionary [--out <csv path>] [--report <md path>]

Both outputs default to <sample dir>/seed_output/ (outside the repo). Order files hold customer
data; the CSV and the report contain product-level fields only. Nothing is uploaded or committed.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd

from core.config_loader import load_config
from core.dictionary import DictionarySettings, ItemDictionary
from core.seed import HandEvidence
from store.csv_codec import from_csv_text, to_csv_text
from tests.regression_harness import order_password, sample_dir
from tools.seed_build import Key, SeedResult, build_seed, dictionary_frame
from tools.seed_data import (
    PoJoin,
    PoPair,
    discover_order_files,
    find_po_pairs,
    join_po_to_orders,
    load_all_orders,
    load_order_frame,
    load_po_frame,
)
from tools.seed_eval import evaluate_pair
from tools.seed_report import PoEval, render_report, totals

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config.xlsx"
OUTPUT_DIRNAME = "seed_output"
PO_FOLDER = Path("Version 2.0") / "Rule 작성을 위한 과거자료"
KST = timezone(timedelta(hours=9))


def _load_pairs(
    pairs: list[PoPair], password: str | None, settings: DictionarySettings
) -> tuple[list[tuple[PoPair, pd.DataFrame, PoJoin]], int]:
    loaded = []
    skipped = 0
    for pair in pairs:
        po = load_po_frame(pair.po_path, password)
        orders = load_order_frame(pair.order_path, password)
        if po is None or orders is None:
            skipped += 1
            continue
        loaded.append((pair, orders, join_po_to_orders(po, orders, settings)))
    return loaded, skipped


def _merge_evidence(joins: list[PoJoin]) -> dict[Key, list[HandEvidence]]:
    merged: dict[Key, list[HandEvidence]] = {}
    for join in joins:
        for key, items in join.evidence.items():
            merged.setdefault(key, []).extend(items)
    return merged


def _print_summary(seed: SeedResult, evals: list[PoEval], pairs_skipped: int) -> None:
    sources = [i.source for i in seed.items]
    total = totals(evals)
    print(f"keys (entries): {len(seed.items)}")
    print(f"  past_po: {sources.count('past_po')}, auto_rule: {sources.count('auto_rule')}")
    print(f"skipped: {sum(seed.skipped.values())} {seed.skipped}")
    print(f"probe failures: {len(seed.probe_failures)}, verification failures: {len(seed.verify_failures)}")
    print(f"qty-1 stray-stamp entries (template from qty 2-3): {seed.q1_differs}")
    print(f"PO pairs skipped (unreadable/encrypted): {pairs_skipped}, bundles evaluated: {total.bundles}")
    print(f"exact match      A {total.exact_a} -> B {total.exact_b}")
    print(f"normalized match A {total.norm_a} -> B {total.norm_b}")


def run(out_path: Path, report_path: Path, separator: str = DictionarySettings().item_separator) -> None:
    password = order_password()
    root = sample_dir()
    config = load_config(str(CONFIG_PATH))
    settings = DictionarySettings(item_separator=separator)

    load = load_all_orders(discover_order_files(root, ROOT, exclude=(root / OUTPUT_DIRNAME,)), password)
    pairs = find_po_pairs(root / PO_FOLDER)
    loaded, pairs_skipped = _load_pairs(pairs, password, settings)

    seed = build_seed(load.orders, config, _merge_evidence([j for _, _, j in loaded]), settings)
    frame = dictionary_frame(seed.items, datetime.now(KST))
    csv_text = to_csv_text(frame)
    dictionary = ItemDictionary(from_csv_text(csv_text, as_text=True), settings)

    evals = [
        PoEval(
            pair.label,
            evaluate_pair(orders, join.bundles, config, dictionary, settings),
            join.po_rows, join.unmatched_rows, join.ambiguous_rows, join.no_hand_bundles,
        )
        for pair, orders, join in loaded
    ]
    out_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(csv_text, encoding="utf-8", newline="")
    report = render_report(seed, load, evals, len(pairs), pairs_skipped, password is not None)
    report_path.write_text(report, encoding="utf-8")
    _print_summary(seed, evals, pairs_skipped)
    print(f"csv: {out_path}")
    print(f"report: {report_path}")


def main(argv: list[str] | None = None) -> None:
    default_dir = sample_dir() / OUTPUT_DIRNAME
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--out", type=Path, default=default_dir / "dictionary_seed.csv", help="draft dictionary CSV")
    parser.add_argument("--report", type=Path, default=default_dir / "seed_report.md", help="markdown report")
    parser.add_argument(
        "--separator", default=DictionarySettings().item_separator,
        help='item separator used when evaluating (settings.json item_separator), e.g. ", "',
    )
    args = parser.parse_args(argv)
    run(args.out, args.report, args.separator)


if __name__ == "__main__":
    main()
