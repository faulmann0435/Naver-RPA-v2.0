"""Migrate the OutputLayout sheet (purchase-order forms) from config.xlsx into the private data repository.

Usage:
    python -m tools.migrate_layout --branch dev --dry-run
    python -m tools.migrate_layout --branch dev
    python -m tools.migrate_layout --branch dev --force     # overwrite output_layout.csv if it differs

Create-only by default: an existing output_layout.csv with other content is never touched without --force.
The token is read from .streamlit/secrets.toml ([data_store]) and is never printed.
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime

import pandas.testing as pdt
import tomllib

from core.config_loader import load_config, normalize_config, read_config_sheets
from store.base import DataStore
from store.layout_repo import OUTPUT_LAYOUT_FILE, layout_csv_to_raw, layout_sheet_to_csv
from tools.migrate_rules import (
    CONFIG_PATH,
    KST,
    SECRETS_PATH,
    apply_plan,
    build_plan,
    load_store,
)


def build_files(config_path: str, migrated_by: str, now: str) -> dict[str, str]:
    """The files to write: output_layout.csv converted from the OutputLayout sheet."""
    sheet = read_config_sheets(config_path)["OutputLayout"]
    return {OUTPUT_LAYOUT_FILE: layout_sheet_to_csv(sheet, migrated_by, now)}


def verify_round_trip(config_path: str, files: dict[str, str]) -> None:
    """Abort (AssertionError) unless the CSV gives the same normalized OutputLayout as config.xlsx."""
    expected = load_config(config_path)["OutputLayout"]
    sheets = read_config_sheets(config_path)
    actual = normalize_config({**sheets, "OutputLayout": layout_csv_to_raw(files[OUTPUT_LAYOUT_FILE])})["OutputLayout"]
    pdt.assert_frame_equal(actual, expected)


def align_with_existing(store: DataStore, files: dict[str, str]) -> dict[str, str]:
    """Keep the stored text when it holds the same layout (only the migration timestamp differs): a re-run is a skip."""
    current = store.read_text(OUTPUT_LAYOUT_FILE)
    if current is not None and layout_csv_to_raw(current.content).equals(layout_csv_to_raw(files[OUTPUT_LAYOUT_FILE])):
        return {OUTPUT_LAYOUT_FILE: current.content}
    return files


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--branch", help="override the branch from secrets (e.g. dev)")
    parser.add_argument("--dry-run", action="store_true", help="print the plan without writing")
    parser.add_argument("--force", action="store_true", help="overwrite output_layout.csv if its content differs")
    args = parser.parse_args(argv)

    now = datetime.now(KST).isoformat(timespec="seconds")
    files = build_files(str(CONFIG_PATH), "migration", now)
    try:
        verify_round_trip(str(CONFIG_PATH), files)
    except AssertionError as e:
        print(f"ABORT: CSV round trip does not reproduce the OutputLayout of config.xlsx:\n{e}")
        return 1
    print("round trip check: OK")

    try:
        store = load_store(args.branch)
    except (OSError, KeyError, tomllib.TOMLDecodeError) as e:
        print(f"ABORT: cannot read [data_store] from {SECRETS_PATH.name}: {type(e).__name__}")
        return 1
    print(f"target: {store.repo}@{store.branch}")
    plan = build_plan(store, align_with_existing(store, files), args.force)
    for item in plan:
        print(f"  {item.filename:<20} rows={item.rows:<4} action={item.action}")
    if any(item.action == "refuse" for item in plan):
        print("output_layout.csv already exists with different content. Re-run with --force to overwrite.")
        if not args.dry_run:
            return 1
    if args.dry_run:
        print("dry run: nothing written")
        return 0
    apply_plan(store, align_with_existing(store, files), plan)
    print("done")
    return 0


if __name__ == "__main__":
    sys.exit(main())
