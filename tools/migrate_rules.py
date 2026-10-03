"""Migrate ProductRoute / OptionRules from config.xlsx into the private data repository.

Usage:
    python -m tools.migrate_rules --branch dev --dry-run
    python -m tools.migrate_rules --branch dev
    python -m tools.migrate_rules --branch dev --force     # overwrite files that differ

The token is read from .streamlit/secrets.toml ([data_store]) and is never printed.
"""
from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import pandas.testing as pdt
import tomllib

from core.config_loader import load_config, normalize_config, read_config_sheets
from core.engine import IMPLEMENTED_ACTIONS
from store.base import DataStore
from store.csv_codec import from_csv_text
from store.github_store import GitHubStore
from store.rules_repo import (
    OPTION_RULES_FILE,
    PRODUCT_ROUTE_FILE,
    rule_csvs_to_raw_sheets,
    sheets_to_rule_csvs,
)

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config.xlsx"
SECRETS_PATH = ROOT / ".streamlit" / "secrets.toml"
KST = timezone(timedelta(hours=9))


@dataclass(frozen=True)
class PlanItem:
    """What will happen to one file."""

    filename: str
    rows: int
    action: str  # create | skip | update | refuse
    sha: str | None = None


def _count_rows(filename: str, text: str) -> int:
    if filename.endswith(".csv"):
        return len(from_csv_text(text))
    return 1


def verify_round_trip(config_path: str, files: dict[str, str]) -> None:
    """Abort (AssertionError) unless the CSV round trip reproduces load_config for the rule sheets."""
    expected = load_config(config_path)
    layout = read_config_sheets(config_path)["OutputLayout"]
    sheets = rule_csvs_to_raw_sheets(files[PRODUCT_ROUTE_FILE], files[OPTION_RULES_FILE])
    actual = normalize_config({**sheets, "OutputLayout": layout})
    pdt.assert_frame_equal(actual["ProductRoute"], expected["ProductRoute"])
    exp_rules: pd.DataFrame = expected["OptionRules"]
    actions = exp_rules["ActionType"].astype(str).str.strip().str.upper()
    kept = exp_rules[actions.isin(IMPLEMENTED_ACTIONS)].reset_index(drop=True)
    pdt.assert_frame_equal(actual["OptionRules"], kept)


def build_plan(store: DataStore, files: dict[str, str], force: bool) -> list[PlanItem]:
    """Decide per file: create / skip (identical) / update (--force) / refuse."""
    plan: list[PlanItem] = []
    for name, text in files.items():
        rows = _count_rows(name, text)
        current = store.read_text(name)
        if current is None:
            plan.append(PlanItem(name, rows, "create"))
        elif current.content == text:
            plan.append(PlanItem(name, rows, "skip", current.sha))
        elif force:
            plan.append(PlanItem(name, rows, "update", current.sha))
        else:
            plan.append(PlanItem(name, rows, "refuse", current.sha))
    return plan


def apply_plan(store: DataStore, files: dict[str, str], plan: list[PlanItem]) -> None:
    """Write create/update items."""
    for item in plan:
        if item.action == "create":
            store.write_text(item.filename, files[item.filename], None, _message(item.filename))
        elif item.action == "update":
            store.write_text(item.filename, files[item.filename], item.sha, _message(item.filename))


def _message(filename: str) -> str:
    return f"chore(data): migrate {filename} from config.xlsx"


def load_store(branch_override: str | None) -> GitHubStore:
    with SECRETS_PATH.open("rb") as f:
        section = tomllib.load(f)["data_store"]
    return GitHubStore(
        repo=section["repo"],
        branch=branch_override or section["branch"],
        token=section["token"],
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--branch", help="override the branch from secrets (e.g. dev)")
    parser.add_argument("--dry-run", action="store_true", help="print the plan without writing")
    parser.add_argument("--force", action="store_true", help="overwrite files whose content differs")
    args = parser.parse_args(argv)

    now = datetime.now(KST).isoformat(timespec="seconds")
    files = sheets_to_rule_csvs(read_config_sheets(str(CONFIG_PATH)), migrated_by="migration", now=now)
    try:
        verify_round_trip(str(CONFIG_PATH), files)
    except AssertionError as e:
        print(f"ABORT: CSV round trip does not reproduce config.xlsx rules:\n{e}")
        return 1
    print("round trip check: OK")

    try:
        store = load_store(args.branch)
    except (OSError, KeyError, tomllib.TOMLDecodeError) as e:
        print(f"ABORT: cannot read [data_store] from {SECRETS_PATH.name}: {type(e).__name__}")
        return 1
    print(f"target: {store.repo}@{store.branch}")
    plan = build_plan(store, files, args.force)
    for item in plan:
        print(f"  {item.filename:<20} rows={item.rows:<4} action={item.action}")
    if any(item.action == "refuse" for item in plan):
        print("Some files already exist with different content. Re-run with --force to overwrite.")
        if not args.dry_run:
            return 1
    if args.dry_run:
        print("dry run: nothing written")
        return 0
    apply_plan(store, files, plan)
    print("done")
    return 0


if __name__ == "__main__":
    sys.exit(main())
