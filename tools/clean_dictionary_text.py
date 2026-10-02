"""
Remove emoji / decorative symbols (except ★☆) and invisible characters from the dictionary's
purchase-order text fields (display_template, display_template_qty1, sum_group).

    python -m tools.clean_dictionary_text --branch dev --dry-run
    python -m tools.clean_dictionary_text --branch dev

Reference fields (product_name_ref, option_raw_ref, option_key) are never changed. The token is read
from .streamlit/secrets.toml ([data_store]) and never printed.
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
import tomllib

from core.template import clean_display_text
from store.base import Author
from store.dictionary_repo import load_dictionary_frame, now_kst_iso, save_dictionary
from store.github_store import GitHubStore

ROOT = Path(__file__).resolve().parent.parent
TEXT_FIELDS = ("display_template", "display_template_qty1", "sum_group")


def cleaned_frame(frame: pd.DataFrame, author: str, now: str) -> tuple[pd.DataFrame, int, int]:
    """(new frame, changed cells, changed rows); changed rows get updated_at / updated_by."""
    out = frame.fillna("").astype(str).copy()
    cells = 0
    changed_rows = set()
    for field in TEXT_FIELDS:
        for i, value in out[field].items():
            new = clean_display_text(value)
            if new != value:
                out.at[i, field] = new
                cells += 1
                changed_rows.add(i)
    for i in changed_rows:
        out.at[i, "updated_at"] = now
        out.at[i, "updated_by"] = author
    return out, cells, len(changed_rows)


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
    author = "clean tool"
    new, cells, rows = cleaned_frame(frame, author, now_kst_iso())
    print(f"target: {store.repo}@{args.branch} | entries {len(frame)} | cells to clean {cells} (rows {rows})")
    if args.dry_run or cells == 0:
        print("dry run: nothing written" if args.dry_run else "nothing to clean")
        return
    message = f"사전 정리: 표기의 이모지·보이지 않는 문자 제거 {cells}칸 ({rows}항목, ★☆ 유지)"
    save_dictionary(store, new, sha, Author(author, "clean@local"), message)
    print("done:", message)


if __name__ == "__main__":
    main()
