"""Link recent product Q&A and 1:1 inquiries to their orders (read-only) and print a masked summary.

    python -m worker.cs_link [--days 3] [--lookback 14] [--out DIR]

The console shows channel, time, answered, link kind and order status only — never question text.
The full result (with question text) goes to CS자료/주문연결_<time>/ next to the repo
(customer data: outside the repo).
"""
from __future__ import annotations

import argparse
import io
import json
import sys
from dataclasses import asdict
from datetime import datetime, timedelta
from pathlib import Path

from worker.inquiry_client import fetch_customer_inquiries, fetch_qnas
from worker.naver_adapter import KST
from worker.naver_client import NaverApiError, fetch_details, fetch_token, load_credentials
from worker.order_link import (
    DEFAULT_LOOKBACK_DAYS,
    LinkResult,
    RetryingSession,
    fetch_recent_details,
    link_customer_inquiry,
    link_qna,
    missing_inquiry_ids,
    parse_time,
)
from worker.paths import DATA_DIR

CS_DIR_NAME = "CS자료"
CHANNEL_QNA, CHANNEL_INQUIRY = "상품문의", "고객문의"
ID_TAIL = 4


def _stamp(moment: datetime | None) -> str:
    return moment.astimezone(KST).strftime("%Y-%m-%d %H:%M") if moment else "?"


def _record(channel: str, inquiry_id: str, created: str | None, answered: bool, text: str,
            result: LinkResult) -> dict:
    return {"channel": channel, "id": inquiry_id, "created": created, "answered": answered,
            "question": text, "link_kind": result.kind, "note": result.note,
            "orders": [_order_dict(asdict(o)) for o in result.orders]}


def _order_dict(order: dict) -> dict:
    return {k: (v.isoformat() if isinstance(v, datetime) else v) for k, v in order.items()}


def build_records(qnas: list[dict], inquiries: list[dict], details: list[dict]) -> list[dict]:
    records = [_record(CHANNEL_QNA, str(q.get("questionId")), q.get("createDate"), bool(q.get("answered")),
                       q.get("question") or "", link_qna(q, details)) for q in qnas]
    records += [_record(CHANNEL_INQUIRY, str(i.get("inquiryNo")), i.get("inquiryRegistrationDateTime"),
                        bool(i.get("answered")), i.get("inquiryContent") or "",
                        link_customer_inquiry(i, details)) for i in inquiries]
    return records


def console_line(record: dict) -> str:
    """One masked line: no question text, order ids cut to their last 4 digits."""
    orders = record["orders"]
    tails = ",".join("…" + o["order_id"][-ID_TAIL:] for o in orders)
    order_part = ""
    if orders:
        first = orders[0]
        order_part = f" | 주문 {tails} 결제 {_stamp(parse_time(first['paid_at']))} {first['status_label']}"
    return (f"[{record['channel']}] {_stamp(parse_time(record['created']))} "
            f"답변 {'예' if record['answered'] else '아니오'} | 연결: {record['link_kind']}{order_part}")


def save(out_dir: Path, records: list[dict]) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "주문연결.json"
    path.write_text(json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8")
    return path


def run(args: argparse.Namespace) -> int:
    now = datetime.now(KST)
    end = now.date()
    start = end - timedelta(days=args.days - 1)
    token = fetch_token(load_credentials())
    qnas = fetch_qnas(token, start, end)
    inquiries = fetch_customer_inquiries(token, start, end)
    details = fetch_recent_details(token, args.lookback, now=now)
    missing = missing_inquiry_ids(inquiries, details)
    if missing:  # 1:1 inquiries name their orders: fetch those directly even when older than the lookback
        details = details + fetch_details(token, missing, RetryingSession())
    records = build_records(qnas, inquiries, details)

    print(f"조회 기간: {start} ~ {end} (주문 조회 {args.lookback}일, 주문 {len(details)}건)")
    for record in records:
        print(console_line(record))
    out_dir = args.out or DATA_DIR / CS_DIR_NAME / f"주문연결_{now:%Y%m%d_%H%M%S}"
    print(f"저장: {save(out_dir, records)}")
    return 0


def main(argv: list[str] | None = None) -> int:
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--days", type=int, default=3, help="문의를 볼 최근 일수")
    parser.add_argument("--lookback", type=int, default=DEFAULT_LOOKBACK_DAYS, help="주문을 찾아볼 최근 일수")
    parser.add_argument("--out", type=Path, default=None, help="저장 폴더 (기본: 저장소 옆 CS자료)")
    args = parser.parse_args(argv)
    try:
        return run(args)
    except NaverApiError as err:
        print(f"네이버 API 오류: {err}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
