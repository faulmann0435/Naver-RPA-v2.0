"""
One-time backfill of the dictionary's created_at: rows with an empty created_at get their updated_at.

    python -m tools.backfill_created_at --branch dev --dry-run
    python -m tools.backfill_created_at --branch dev

updated_at / updated_by are never changed. Rows that already have a created_at are kept. The token is
read from .streamlit/secrets.toml ([data_store]) and never printed.
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
import tomllib

from store.base import Author
from store.dictionary_repo import load_dictionary_frame, save_dictionary
from store.github_store import GitHubStore

ROOT = Path(__file__).resolve().parent.parent


def backfilled_frame(frame: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """(new frame, filled rows): an empty created_at becomes the row's updated_at (still empty if that is empty too)."""
    out = frame.copy()
    if "created_at" not in out.columns:
        out["created_at"] = ""
    out["created_at"] = out["created_at"].fillna("").astype(str)
    updated = out["updated_at"].fillna("").astype(str) if "updated_at" in out.columns else out["created_at"]
    fill = (out["created_at"].str.strip() == "") & (updated.str.strip() != "")
    out.loc[fill, "created_at"] = updated[fill]
    return out, int(fill.sum())


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--branch", required=True, help="data repo branch (dev / main)")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)
    with (ROOT / ".streamlit" / "secrets.toml").open("rb") as f:
        section = dict(tomllib.load(f)["data_store"])
    section["branch"] = args.branch
    store = GitHubStore.from_secrets(section)
    frame, sha = load_dictionary_frame(store)
    new, filled = backfilled_frame(frame)
    print(f"target: {store.repo}@{args.branch} | entries {len(frame)} | rows to fill {filled}")
    if args.dry_run or filled == 0:
        print("dry run: nothing written" if args.dry_run else "nothing to fill")
        return
    message = f"등록 시각 채우기: {filled}건"
    save_dictionary(store, new, sha, Author("backfill tool", "backfill@local"), message)
    print("done:", message)


if __name__ == "__main__":
    main()
