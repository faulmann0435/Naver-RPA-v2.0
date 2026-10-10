"""Read recent paid orders from the real API and report their field formats with customer data masked.

    python -m worker.probe [--days 3]

Read-only. Saves the raw response, the converted table and the purchase orders under
API변환_데모/실제조회_<time>/ (customer data: outside the repo).
"""
from __future__ import annotations

import argparse
import io
import json
import re
import sys
from collections import Counter
from datetime import datetime, timedelta

from core.config_loader import load_config
from core.pipeline import process_all_data
from tests.regression_harness import CONFIG_PATH, DEMO_DIR_NAME, REPO_ROOT
from worker.naver_adapter import KST, details_to_dataframe, split_shippable
from worker.naver_client import (
    fetch_changed_ids,
    fetch_details,
    fetch_token,
    load_credentials,
)


def shape(value) -> str:
    """Digits -> 9, letters -> a, Hangul -> 가: shows a value's format, never its content."""
    text = re.sub(r"\d", "9", str(value))
    text = re.sub(r"[A-Za-z]", "a", text)
    return re.sub(r"[가-힣]+", "가", text)


def _field(item: dict, *path: str):
    current: object = item
    for key in path:
        current = current.get(key) if isinstance(current, dict) else None
    return current


def report(items: list[dict]) -> list[str]:
    """Masked observations about the fields the adapter relies on."""
    count = lambda *path: Counter(str(_field(i, *path)) for i in items)
    shapes = lambda *path: Counter(shape(_field(i, *path)) for i in items)
    addr = [_field(i, "productOrder", "shippingAddress") or {} for i in items]
    return [
        f"주문 상태(productOrderStatus): {dict(count('productOrder', 'productOrderStatus'))}",
        f"발주확인 상태(placeOrderStatus): {dict(count('productOrder', 'placeOrderStatus'))}",
        f"상품종류(productClass): {dict(count('productOrder', 'productClass'))}",
        f"수취인 연락처1 형식: {dict(shapes('productOrder', 'shippingAddress', 'tel1'))}",
        f"구매자 연락처 형식: {dict(shapes('order', 'ordererTel'))}",
        f"우편번호 형식: {dict(shapes('productOrder', 'shippingAddress', 'zipCode'))}",
        f"결제일 형식: {dict(shapes('order', 'paymentDate'))}",
        (f"상세주소 있음: {sum(1 for a in addr if a.get('detailedAddress'))}/{len(addr)}, "
         f"주소 필드: {sorted({k for a in addr for k in a})}"),
        f"배송비 묶음번호(packageNumber) 있음: {sum(1 for i in items if _field(i, 'productOrder', 'packageNumber'))}/{len(items)}",
    ]


def main(argv: list[str] | None = None) -> int:
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--days", type=int, default=3)
    args = parser.parse_args(argv)

    now = datetime.now(KST)
    token = fetch_token(load_credentials())
    ids = fetch_changed_ids(token, now - timedelta(days=args.days), now)
    items = fetch_details(token, ids)
    print(f"최근 {args.days}일 결제완료로 바뀐 상품주문 {len(ids)}건, 상세 {len(items)}건")
    if not items:
        return 0
    for line in report(items):
        print(f"  · {line}")

    out_dir = REPO_ROOT.parent / DEMO_DIR_NAME / f"실제조회_{now:%Y%m%d_%H%M%S}"
    (out_dir / "발주서").mkdir(parents=True, exist_ok=True)
    (out_dir / "네이버API_응답.json").write_text(json.dumps({"data": items}, ensure_ascii=False, indent=2),
                                             encoding="utf-8")
    shippable, left_out = split_shippable(items)
    print(f"발주 대상 {len(shippable)}건, 제외 {len(left_out)}건 (취소·반품·교환 요청 또는 이미 발송)")
    table = details_to_dataframe(shippable)
    table.to_excel(out_dir / "변환된_주문표.xlsx", index=False)
    files = process_all_data(table, load_config(str(CONFIG_PATH)))
    for f in files:
        (out_dir / "발주서" / f["filename"]).write_bytes(f["data"].getvalue())
    print(f"변환 {len(table)}줄 → 발주서 {len(files)}개: {[f['vendor'] for f in files]}")
    print(f"결과 폴더: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
