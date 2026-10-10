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
from dataclasses import dataclass, field
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
# Columns that may differ without anything being wrong (the order moved on since the download).
EXPECTED_DIFFS = frozenset({"주문상태"})
# Naver masks buyer/recipient details of older, finished orders with '*' (e.g. 010-1***-5***).
PERSONAL_COLUMNS = ("구매자명", "수취인명", "구매자연락처", "수취인연락처1", "수취인연락처2", "통합배송지")


def masked_rows(api: pd.DataFrame) -> int:
    """Rows where the API sent personal details masked (older, finished orders)."""
    present = [c for c in PERSONAL_COLUMNS if c in api.columns]
    return int(api[present].astype(str).apply(lambda col: col.str.contains("*", regex=False)).any(axis=1).sum())


def column_differences(excel: pd.DataFrame, api: pd.DataFrame) -> dict[str, list[str]]:
    """Per column, masked examples of rows (matched by 상품주문번호) whose values differ."""
    left = excel.drop_duplicates(KEY).set_index(KEY).sort_index()
    right = api.drop_duplicates(KEY).set_index(KEY).reindex(left.index)
    found: dict[str, list[str]] = {}
    for column in (c for c in COLUMN_ORDER if c != KEY and c in left.columns):
        a, b = left[column], right[column]
        differs = ~((a == b) | (a.isna() & b.isna()))
        if differs.any():
            found[column] = [f"엑셀 {shape(x)!r} / API {shape(y)!r}"
                             for x, y in list(zip(a[differs], b[differs]))[:EXAMPLES]] + [f"({int(differs.sum())}줄)"]
    return found


@dataclass(frozen=True)
class CompareResult:
    """Outcome of lining one Excel export up against the API."""

    file: str
    excel_rows: int = 0
    api_rows: int = 0
    duplicates: int = 0
    masked: int = 0
    same_columns: tuple[str, ...] = ()
    column_diffs: dict[str, list[str]] = field(default_factory=dict)
    file_problems: tuple[str, ...] = ()
    error: str = ""

    @property
    def matched(self) -> bool:
        unexpected = set(self.column_diffs) - EXPECTED_DIFFS
        return not self.error and not self.file_problems and not unexpected and self.excel_rows == self.api_rows

    def summary(self) -> str:
        """One line for a notification."""
        if self.error:
            return f"⚠ {self.file}: 대조 못 함 — {self.error}"
        if self.matched:
            return f"✅ {self.file}: {self.excel_rows}줄, 발주서 모두 일치"
        what = list(self.file_problems) or [f"다른 칸 {sorted(set(self.column_diffs) - EXPECTED_DIFFS)}"]
        return f"❌ {self.file}: 차이 있음 — " + "; ".join(what)

    def details(self) -> list[str]:
        """Full report, customer values masked."""
        if self.error:
            return [self.summary()]
        lines = [f"엑셀 {self.excel_rows}줄, API에서 받은 {self.api_rows}줄"
                 + (f" (엑셀에 같은 상품주문번호가 {self.duplicates}줄 중복)" if self.duplicates else "")]
        if self.masked:
            lines.append(f"⚠ {self.masked}줄은 네이버가 개인정보를 가려서 보냄(오래된·완료된 주문) — 이름·연락처·주소 칸 차이는 정상")
        lines.append(f"같은 칸 {len(self.same_columns)}개: {list(self.same_columns)}")
        lines += [f"다른 칸 '{c}': " + "; ".join(ex) for c, ex in self.column_diffs.items()]
        lines.append("발주서 비교: " + ("모두 같음 ✅" if not self.file_problems else "; ".join(self.file_problems)))
        return lines


def compare_file(path: Path, token: str | None = None) -> CompareResult:
    """Read the export, fetch the same orders from the API and compare tables and purchase orders."""
    excel = filter_instruction_rows(_read_order(load_excel, path)[0])
    if KEY not in excel.columns:
        return CompareResult(path.name, error=f"'{KEY}' 칸이 없는 파일")
    ids = excel[KEY].drop_duplicates().astype(str)
    api = details_to_dataframe(fetch_details(token or fetch_token(load_credentials()), ids))
    diffs = column_differences(excel, api)
    config = load_config(str(CONFIG_PATH))
    return CompareResult(
        file=path.name,
        excel_rows=len(excel.drop_duplicates(KEY)),
        api_rows=len(api),
        duplicates=int(excel[KEY].duplicated().sum()),
        masked=masked_rows(api),
        same_columns=tuple(c for c in COLUMN_ORDER if c in excel.columns and c not in diffs),
        column_diffs=diffs,
        file_problems=tuple(_compare_files(process_all_data(excel, config), process_all_data(api, config))),
    )


def main(argv: list[str] | None = None) -> int:
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("order_file", type=Path)
    args = parser.parse_args(argv)
    result = compare_file(args.order_file)
    for line in result.details():
        print(line)
    return 0 if result.matched else 1


if __name__ == "__main__":
    raise SystemExit(main())
