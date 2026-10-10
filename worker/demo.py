"""Show the API path on a real past order file, step by step, and compare it with the Excel path.

    python -m worker.demo "<order export .xlsx>" [--out <folder>]

Writes the simulated API response, the converted order table and the purchase orders to the
output folder (default: API변환_데모 next to the repo — it holds customer data, keep it out of git).
"""
from __future__ import annotations

import argparse
import io
import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

from core.config_loader import load_config
from core.loader import load_excel
from core.merger import filter_instruction_rows
from core.pipeline import process_all_data
from tests.regression_harness import (
    CONFIG_PATH,
    DEMO_DIR_NAME,
    REPO_ROOT,
    _cells,
    _read_order,
)
from worker.naver_adapter import COLUMN_ORDER, KST, details_to_dataframe
from worker.simulator import excel_to_api_items

DEFAULT_OUT = REPO_ROOT.parent / DEMO_DIR_NAME


def _same_columns(excel_df: pd.DataFrame, api_df: pd.DataFrame) -> list[str]:
    """Columns the adapter fills whose values differ from the Excel export."""
    shared = [c for c in COLUMN_ORDER if c in excel_df.columns]
    return [c for c in shared if not excel_df[c].reset_index(drop=True).equals(api_df[c].reset_index(drop=True))]


def _compare_files(excel_files: list[dict], api_files: list[dict]) -> list[str]:
    by_vendor = {f["vendor"]: f for f in excel_files}
    problems = []
    for f in api_files:
        other = by_vendor.get(f["vendor"])
        if other is None:
            problems.append(f"{f['vendor']}: 엑셀 경로에는 없는 파일")
        elif not _cells(f["data"].getvalue()).astype(str).equals(_cells(other["data"].getvalue()).astype(str)):
            problems.append(f"{f['vendor']}: 내용이 다름")
    missing = set(by_vendor) - {f["vendor"] for f in api_files}
    return problems + [f"{v}: API 경로에서 빠짐" for v in sorted(missing)]


def run(order_path: Path, out_root: Path) -> int:
    out_dir = out_root / datetime.now(KST).strftime("%Y%m%d_%H%M%S")
    (out_dir / "3_발주서").mkdir(parents=True, exist_ok=True)
    config = load_config(str(CONFIG_PATH))

    excel_df = filter_instruction_rows(_read_order(load_excel, order_path)[0])
    print(f"\n주문 파일: {order_path.name} — 주문 {len(excel_df)}줄")

    items = excel_to_api_items(excel_df)
    json_path = out_dir / "1_네이버API_응답(모의).json"
    json_path.write_text(json.dumps({"data": items}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"① 네이버 API가 보내줄 모양(JSON)으로 바꿈 → {json_path.name}")

    api_df = details_to_dataframe(items)
    api_df.to_excel(out_dir / "2_변환된_주문표.xlsx", index=False)
    diff_cols = _same_columns(excel_df, api_df)
    print(f"② 변환기가 엑셀과 같은 표로 만듦 → 칸 {len(api_df.columns)}개, "
          + ("엑셀과 모두 같음" if not diff_cols else f"다른 칸: {diff_cols}"))

    api_files = process_all_data(api_df, config)
    for f in api_files:
        (out_dir / "3_발주서" / f["filename"]).write_bytes(f["data"].getvalue())
    print(f"③ v15 엔진 실행 → 발주서 {len(api_files)}개")
    for f in api_files:
        print(f"     · {f['filename']} ({len(_cells(f['data'].getvalue()))}줄)")

    problems = _compare_files(process_all_data(excel_df, config), api_files)
    print("④ 지금 방식(엑셀 업로드) 결과와 비교 → "
          + ("모든 발주서가 한 글자도 다르지 않음 ✅" if not problems else "차이 있음 ❌"))
    for p in problems:
        print(f"     · {p}")
    print(f"\n결과 폴더: {out_dir}")
    return 0 if not problems and not diff_cols else 1


def main(argv: list[str] | None = None) -> int:
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("order_file", type=Path)
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT)
    args = parser.parse_args(argv)
    if not args.order_file.exists():
        print(f"파일이 없습니다: {args.order_file}")
        return 2
    return run(args.order_file.resolve(), args.out)


if __name__ == "__main__":
    raise SystemExit(main())
