"""Collect past product Q&A and 1:1 inquiries from the real API and print a summary without their text.

    python -m worker.cs_collect [--days 365]

Read-only. Saves the raw lists under CS자료/문의수집_<time>/ next to the repo
(customer data: outside the repo). The summary shows counts only, never question text or names.
"""
from __future__ import annotations

import argparse
import io
import json
import sys
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

from tests.regression_harness import REPO_ROOT
from worker.inquiry_client import fetch_customer_inquiries, fetch_qnas
from worker.naver_adapter import KST
from worker.naver_client import fetch_token, load_credentials

CS_DIR_NAME = "CS자료"
TOP_PRODUCTS = 10


def _month(text: str | None) -> str:
    return (text or "")[:7] or "?"


def summarize_qnas(qnas: list[dict]) -> list[str]:
    months = Counter(_month(q.get("createDate")) for q in qnas)
    products = Counter(q.get("productName") or "?" for q in qnas)
    answered = sum(1 for q in qnas if q.get("answered"))
    lines = [f"상품문의 {len(qnas)}건 (답변 완료 {answered}건, 미답변 {len(qnas) - answered}건)",
             f"  월별: {dict(sorted(months.items()))}",
             f"  많이 묻는 상품 TOP{TOP_PRODUCTS}:"]
    return lines + [f"    {count:>4}건  {name}" for name, count in products.most_common(TOP_PRODUCTS)]


def summarize_inquiries(inquiries: list[dict]) -> list[str]:
    months = Counter(_month(i.get("inquiryRegistrationDateTime")) for i in inquiries)
    categories = Counter(i.get("category") or "?" for i in inquiries)
    answered = sum(1 for i in inquiries if i.get("answered"))
    with_order = sum(1 for i in inquiries if i.get("orderId"))
    return [f"고객문의 {len(inquiries)}건 (답변 완료 {answered}건, 주문 연결 {with_order}건)",
            f"  월별: {dict(sorted(months.items()))}",
            f"  문의 유형(category): {dict(categories.most_common())}"]


def save(out_dir: Path, qnas: list[dict], inquiries: list[dict]) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for name, rows in (("상품문의.json", qnas), ("고객문의.json", inquiries)):
        (out_dir / name).write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--days", type=int, default=365)
    args = parser.parse_args(argv)

    now = datetime.now(KST)
    end = now.date()
    start = end - timedelta(days=args.days - 1)
    token = fetch_token(load_credentials())
    qnas = fetch_qnas(token, start, end)
    inquiries = fetch_customer_inquiries(token, start, end)

    print(f"조회 기간: {start} ~ {end}")
    for line in summarize_qnas(qnas) + summarize_inquiries(inquiries):
        print(line)
    out_dir = REPO_ROOT.parent / CS_DIR_NAME / f"문의수집_{now:%Y%m%d_%H%M%S}"
    save(out_dir, qnas, inquiries)
    print(f"원본 저장: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
