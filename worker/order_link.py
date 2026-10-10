"""Read-only link from a customer question (1:1 inquiry or product Q&A) to the order it is about.

Only non-personal order facts are kept (status, dates, product, courier) — never names, phones,
addresses, the orderer id or tracking numbers. Nothing here changes anything on the store.
"""
from __future__ import annotations

import re
import time
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime, timedelta

import requests

from worker.inquiry_client import RATE_LIMIT_STATUS, RETRY_WAITS
from worker.naver_adapter import KST
from worker.naver_client import fetch_changed_ids, fetch_details

STATUS_LABELS = {
    "PAYED": "결제완료(발송 전)",
    "DELIVERING": "배송중",
    "DELIVERED": "배송완료",
    "PURCHASE_DECIDED": "구매확정",
    "CANCELED": "취소",
    "RETURNED": "반품",
    "EXCHANGED": "교환",
    "CANCELED_BY_NOPAYMENT": "미입금 취소",
    "PAYMENT_WAITING": "입금대기",
}
KIND_EXACT, KIND_GUESS, KIND_MANY, KIND_NONE = "정확", "추정", "후보 여러 건", "없음"
MIN_VISIBLE_ID_CHARS = 2  # a shorter visible id prefix says too little to match on
MAX_CANDIDATES = 3
DEFAULT_LOOKBACK_DAYS = 14
NOTE_NOT_FOUND = "주문 정보를 찾지 못했습니다"
NOTE_GUESS = "작성자 아이디 앞부분·상품·결제 시각으로 추정한 주문입니다"
NOTE_MANY = "조건에 맞는 주문이 여러 건이라 최근 {n}건만 후보로 보여줍니다"
NOTE_EXACT = "문의에 적힌 주문번호로 연결했습니다"


class RetryingSession:
    """Wraps a requests-like session: GET/POST are retried after a 429 rate-limit refusal."""

    def __init__(self, inner: requests.Session | None = None,
                 sleep: Callable[[float], None] = time.sleep) -> None:
        self._inner = inner or requests.Session()
        self._sleep = sleep

    def _call(self, method: Callable[..., requests.Response], url: str, **kwargs) -> requests.Response:
        resp = method(url, **kwargs)
        for wait in RETRY_WAITS:
            if resp.status_code != RATE_LIMIT_STATUS:
                break
            self._sleep(wait)
            resp = method(url, **kwargs)
        return resp

    def get(self, url: str, **kwargs) -> requests.Response:
        return self._call(self._inner.get, url, **kwargs)

    def post(self, url: str, **kwargs) -> requests.Response:
        return self._call(self._inner.post, url, **kwargs)


@dataclass(frozen=True)
class OrderSummary:
    order_id: str
    product_order_ids: tuple[str, ...]
    product_name: str
    option: str
    quantity: int
    status: str
    status_label: str
    paid_at: datetime | None
    shipping_due: datetime | None
    arrival_guarantee: datetime | None
    courier: str | None
    sent_at: datetime | None
    delivered_at: datetime | None
    memo: str


@dataclass(frozen=True)
class LinkResult:
    kind: str
    orders: tuple[OrderSummary, ...]
    note: str


def parse_time(text: str | None) -> datetime | None:
    """ISO text from the API as a KST-aware datetime; None when empty or unreadable."""
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(str(text).replace("Z", "+00:00"))
    except ValueError:
        return None
    return parsed.replace(tzinfo=KST) if parsed.tzinfo is None else parsed.astimezone(KST)


LONG_NUMBER = re.compile(r"[\d*#][\d\-*# ]{2,}[\d*#]")  # phones, door codes, unit numbers
MASK = "[숫자]"


def clean_memo(memo: str | None) -> str:
    """Delivery memo with phone numbers and door codes hidden (it can go to the AI drafter)."""
    return LONG_NUMBER.sub(MASK, memo or "").strip()


def _unique(values) -> list[str]:
    return list(dict.fromkeys(v for v in values if v))


def _first_time(parts: list[dict], key: str) -> datetime | None:
    return next((t for t in (parse_time(p.get(key)) for p in parts) if t), None)


def _summary(order_id: str, items: list[dict]) -> OrderSummary:
    orders = [i.get("order") or {} for i in items]
    pos = [i.get("productOrder") or {} for i in items]
    deliveries = [i.get("delivery") or {} for i in items]
    statuses = _unique(p.get("productOrderStatus") for p in pos)
    couriers = _unique(d.get("deliveryCompany") for d in deliveries) or \
        _unique(p.get("expectedDeliveryCompany") for p in pos)
    return OrderSummary(
        order_id=order_id,
        product_order_ids=tuple(_unique(str(p.get("productOrderId")) for p in pos if p.get("productOrderId"))),
        product_name=" / ".join(_unique(p.get("productName") for p in pos)),
        option=" / ".join(_unique(p.get("productOption") for p in pos)),
        quantity=sum(int(p.get("quantity") or 0) for p in pos),
        status=" / ".join(statuses),
        status_label=" / ".join(STATUS_LABELS.get(s, s) for s in statuses),
        paid_at=_first_time(orders, "paymentDate"),
        shipping_due=_first_time(pos, "shippingDueDate"),
        arrival_guarantee=_first_time(pos, "arrivalGuaranteeDate"),
        courier=couriers[0] if couriers else None,
        sent_at=_first_time(deliveries, "sendDate"),
        delivered_at=_first_time(deliveries, "deliveredDate"),
        memo=" / ".join(_unique(clean_memo(p.get("shippingMemo")) for p in pos)),
    )


def summarize(details: list[dict]) -> list[OrderSummary]:
    """One OrderSummary per order id, in order of first appearance."""
    groups: dict[str, list[dict]] = {}
    for item in details:
        order_id = str((item.get("order") or {}).get("orderId") or "")
        if order_id:
            groups.setdefault(order_id, []).append(item)
    return [_summary(order_id, items) for order_id, items in groups.items()]


def _inquiry_po_ids(inquiry: dict) -> set[str]:
    return {s.strip() for s in str(inquiry.get("productOrderIdList") or "").split(",") if s.strip()}


def missing_inquiry_ids(inquiries: list[dict], details: list[dict]) -> list[str]:
    """Product order ids named by 1:1 inquiries but absent from `details` (e.g. older than the lookback)."""
    have = {str((d.get("productOrder") or {}).get("productOrderId")) for d in details}
    return sorted({i for inq in inquiries for i in _inquiry_po_ids(inq)} - have)


def link_customer_inquiry(inquiry: dict, details: list[dict]) -> LinkResult:
    """Link a 1:1 inquiry by its product order ids (fallback: its order id)."""
    wanted = _inquiry_po_ids(inquiry)
    found =[d for d in details if str((d.get("productOrder") or {}).get("productOrderId")) in wanted]
    if not found and inquiry.get("orderId"):
        order_id = str(inquiry["orderId"])
        found = [d for d in details if str((d.get("order") or {}).get("orderId")) == order_id]
    if not found:
        return LinkResult(KIND_NONE, (), NOTE_NOT_FOUND)
    return LinkResult(KIND_EXACT, tuple(summarize(found)), NOTE_EXACT)


def _visible_prefix(masked: str | None) -> str:
    return (masked or "").split("*", 1)[0]


def _qna_matches(qna: dict, item: dict) -> bool:
    prefix = _visible_prefix(qna.get("maskedWriterId"))
    order, po = item.get("order") or {}, item.get("productOrder") or {}
    orderer = order.get("ordererId") or ""
    paid, asked = parse_time(order.get("paymentDate")), parse_time(qna.get("createDate"))
    return (len(prefix) >= MIN_VISIBLE_ID_CHARS
            and _visible_prefix(orderer) == prefix
            and len(orderer) == len(qna.get("maskedWriterId") or "")
            and str(po.get("productId")) == str(qna.get("productId"))
            and paid is not None and asked is not None and paid <= asked)


def link_qna(qna: dict, details: list[dict]) -> LinkResult:
    """Guess the order behind a product Q&A from the masked writer id, product and payment time."""
    summaries = summarize([d for d in details if _qna_matches(qna, d)])
    if not summaries:
        return LinkResult(KIND_NONE, (), NOTE_NOT_FOUND)
    if len(summaries) == 1:
        return LinkResult(KIND_GUESS, tuple(summaries), NOTE_GUESS)
    epoch = datetime.min.replace(tzinfo=KST)
    newest = sorted(summaries, key=lambda s: s.paid_at or epoch, reverse=True)[:MAX_CANDIDATES]
    return LinkResult(KIND_MANY, tuple(newest), NOTE_MANY.format(n=len(newest)))


def fetch_recent_details(token: str, days: int = DEFAULT_LOOKBACK_DAYS,
                         session: requests.Session | None = None,
                         now: datetime | None = None) -> list[dict]:
    """Details of every product order whose last change fell in the last `days` days (any status)."""
    http = RetryingSession(session)
    end = now or datetime.now(KST)
    ids = fetch_changed_ids(token, end - timedelta(days=days), end, None, http)
    return fetch_details(token, ids, http) if ids else []


def _stamp(moment: datetime | None) -> str:
    return moment.astimezone(KST).strftime("%Y-%m-%d %H:%M") if moment else ""


def _order_line(order: OrderSummary) -> str:
    name = order.product_name + (f" ({order.option})" if order.option else "")
    sent = " ".join(p for p in (_stamp(order.sent_at), order.courier or "") if p)
    parts = [f"결제 {_stamp(order.paid_at)}" if order.paid_at else "",
             order.status_label,
             f"{name} ×{order.quantity}",
             f"출고기한 {_stamp(order.shipping_due)}" if order.shipping_due else "",
             f"발송 {sent}" if order.sent_at else (f"택배사 {order.courier}" if order.courier else ""),
             f"도착 {_stamp(order.delivered_at)}" if order.delivered_at else "",
             f"도착보장 {_stamp(order.arrival_guarantee)}" if order.arrival_guarantee else "",
             f"배송메모: {order.memo}" if order.memo else ""]
    return "- " + " / ".join(p for p in parts if p)


def to_prompt_text(result: LinkResult) -> str:
    """Short Korean block for the AI drafter."""
    if result.kind == KIND_NONE:
        return "주문 연결: 없음 (주문 정보를 모르므로 출고일·도착일을 단정하지 말 것)"
    lines = [f"주문 연결: {result.kind}"] + [_order_line(o) for o in result.orders]
    if result.kind == KIND_GUESS:
        lines.append("(추정 연결 — 날짜를 단정하지 말고 사장님 확인용으로만 사용)")
    elif result.kind == KIND_MANY:
        lines.append("(후보가 여러 건 — 어느 주문인지 단정하지 말고 사장님 확인용으로만 사용)")
    return "\n".join(lines)
