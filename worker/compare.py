"""Compare an order Excel export with the same orders read from the real API.

    python -m worker.compare "<order export .xlsx>"

Reads the export's product order ids from the API (read-only), converts them with the adapter
and reports, column by column, where the API table differs from the Excel — values masked —
then compares the purchase orders both paths produce. Encrypted files need RPA_ORDER_PASSWORD.
"""
from __future__ import annotations

import argparse
import io
import sys
from pathlib import Path

import pandas as pd

from core.config_loader import load_config
from core.loader import load_excel
from core.merger import filter_instruction_rows
from core.pipeline import process_all_data
from tests.regression_harness import CONFIG_PATH, _read_order
from worker.demo import _compare_files
from worker.naver_adapter import COLUMN_ORDER, details_to_dataframe
from worker.naver_client import fetch_details, fetch_token, load_credentials
from worker.probe import shape

KEY = "상품주문번호"
EXAMPLES = 3


def column_differences(excel: pd.DataFrame, api: pd.DataFrame) -> dict[str, list[str]]:
    """Per column, masked examples of rows (matched by 상품주문번호) whose values differ."""
    left = excel.set_index(KEY).sort_index()
    right = api.set_index(KEY).reindex(left.index)
    found: dict[str, list[str]] = {}
    for column in (c for c in COLUMN_ORDER if c != KEY and c in left.columns):
        a, b = left[column], right[column]
        differs = ~((a == b) | (a.isna() & b.isna()))
        if differs.any():
            found[column] = [f"엑셀 {shape(x)!r} / API {shape(y)!r}"
                             for x, y in list(zip(a[differs], b[differs]))[:EXAMPLES]] + [f"({int(differs.sum())}줄)"]
    return found


def main(argv: list[str] | None = None) -> int:
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("order_file", type=Path)
    args = parser.parse_args(argv)

    excel = filter_instruction_rows(_read_order(load_excel, args.order_file)[0])
    token = fetch_token(load_credentials())
    api = details_to_dataframe(fetch_details(token, excel[KEY].astype(str)))
    print(f"엑셀 {len(excel)}줄, API에서 받은 {len(api)}줄")

    diffs = column_differences(excel, api)
    same = [c for c in COLUMN_ORDER if c in excel.columns and c not in diffs]
    print(f"같은 칸 {len(same)}개: {same}")
    for column, examples in diffs.items():
        print(f"다른 칸 '{column}': " + "; ".join(examples))

    config = load_config(str(CONFIG_PATH))
    problems = _compare_files(process_all_data(excel, config), process_all_data(api, config))
    print("발주서 비교: " + ("모두 같음 ✅" if not problems else "; ".join(problems)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
