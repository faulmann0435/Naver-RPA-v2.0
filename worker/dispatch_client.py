"""Register tracking numbers on Naver ("발송처리"): the only code that WRITES to the store.

    POST /v1/pay-order/seller/product-orders/dispatch   (30 product orders per call at most)

Safety switch: nothing is sent unless INVOICE_DISPATCH_ENABLED=1 is set in the environment or
.env. Without it dispatch_or_simulate() only pretends (simulated=True) and makes no HTTP call.
"""
from __future__ import annotations

import os
import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from datetime import datetime, timedelta

import requests

from worker.naver_adapter import KST
from worker.naver_client import (
    API_BASE,
    ENV_FILE,
    TIMEOUT_SECONDS,
    _api_time,
    _check,
    fetch_token,
    load_credentials,
    read_env_file,
)
from worker.order_link import RetryingSession

DISPATCH_URL = f"{API_BASE}/v1/pay-order/seller/product-orders/dispatch"
SWITCH_KEY = "INVOICE_DISPATCH_ENABLED"
DELIVERY_METHOD = "DELIVERY"
COURIER_CODE = "HANJIN"
MAX_PER_CALL = 30
MAX_DATE_AGE = timedelta(days=30)  # Naver refuses dispatch dates older than this
TRACKING_PATTERN = re.compile(r"[A-Za-z0-9-]{1,50}")
CODE_REQUEST_FAILED, CODE_NO_RESULT = "REQUEST_FAILED", "NO_RESULT"


@dataclass(frozen=True)
class DispatchItem:
    product_order_id: str
    tracking: str


@dataclass(frozen=True)
class DispatchFailure:
    id: str
    code: str
    message: str


@dataclass(frozen=True)
class DispatchResult:
    success_ids: tuple[str, ...]
    failures: tuple[DispatchFailure, ...] = ()
    simulated: bool = False


def dispatch_enabled() -> bool:
    value = os.environ.get(SWITCH_KEY) or read_env_file(ENV_FILE).get(SWITCH_KEY, "")
    return value.strip() == "1"


def validate_items(items: Sequence[DispatchItem], when: datetime, now: datetime) -> None:
    """Refuse bad input before anything is sent."""
    if when > now:
        raise ValueError("발송일이 미래입니다.")
    if when < now - MAX_DATE_AGE:
        raise ValueError("발송일이 30일보다 오래됐습니다.")
    for item in items:
        if not item.product_order_id:
            raise ValueError("상품주문번호가 비어 있습니다.")
        if not TRACKING_PATTERN.fullmatch(item.tracking):
            raise ValueError(f"송장번호 형식이 올바르지 않습니다 (영문·숫자·하이픈, 50자 이하): 상품주문번호 {item.product_order_id}")


def _body(chunk: Sequence[DispatchItem], when: datetime) -> dict:
    stamp = _api_time(when)
    return {"dispatchProductOrders": [
        {"productOrderId": i.product_order_id, "deliveryMethod": DELIVERY_METHOD,
         "deliveryCompanyCode": COURIER_CODE, "trackingNumber": i.tracking, "dispatchDate": stamp}
        for i in chunk]}


def _read_chunk(response: dict, chunk: Sequence[DispatchItem]) -> tuple[list[str], list[DispatchFailure]]:
    data = response.get("data") or {}
    success = [str(i) for i in data.get("successProductOrderIds") or []]
    failures = [DispatchFailure(str(f.get("productOrderId", "")), str(f.get("code", "")), str(f.get("message", "")))
                for f in data.get("failProductOrderInfos") or []]
    answered = set(success) | {f.id for f in failures}
    failures += [DispatchFailure(i.product_order_id, CODE_NO_RESULT, "네이버 응답에 결과가 없습니다.")
                 for i in chunk if i.product_order_id not in answered]
    return success, failures


def send_dispatch(items: Sequence[DispatchItem], when: datetime | None = None,
                  session: requests.Session | None = None, token: str | None = None,
                  now: Callable[[], datetime] = lambda: datetime.now(KST)) -> DispatchResult:
    """Really send. A failing call keeps what earlier calls achieved and fails only the rest."""
    current = now()
    moment = when or current
    validate_items(items, moment, current)
    if not items:
        return DispatchResult(())
    http = RetryingSession(session)
    headers = {"Authorization": f"Bearer {token or fetch_token(load_credentials(), http)}"}
    success: list[str] = []
    failures: list[DispatchFailure] = []
    for start in range(0, len(items), MAX_PER_CALL):
        chunk = items[start:start + MAX_PER_CALL]
        try:
            response = _check(http.post(DISPATCH_URL, headers=headers, json=_body(chunk, moment),
                                        timeout=TIMEOUT_SECONDS))
        except Exception as exc:  # noqa: BLE001  # keep the record of earlier chunks; fail only the rest
            failures += [DispatchFailure(i.product_order_id, CODE_REQUEST_FAILED, str(exc)[:300])
                         for i in items[start:]]
            break
        chunk_success, chunk_failures = _read_chunk(response, chunk)
        success += chunk_success
        failures += chunk_failures
    return DispatchResult(tuple(success), tuple(failures))


def dispatch_or_simulate(items: Sequence[DispatchItem], enabled: bool | None = None,
                         **kwargs) -> DispatchResult:
    """send_dispatch when the switch is on; otherwise validate and report what would have been sent."""
    on = dispatch_enabled() if enabled is None else enabled
    if on:
        return send_dispatch(items, **kwargs)
    current = kwargs.get("now", lambda: datetime.now(KST))()
    validate_items(items, kwargs.get("when") or current, current)
    return DispatchResult(tuple(i.product_order_id for i in items), (), simulated=True)
