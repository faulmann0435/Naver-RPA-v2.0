"""Read-only Naver Commerce API calls for customer questions: product Q&A and 1:1 inquiries.

Nothing here posts an answer. Token comes from worker.naver_client.fetch_token.
"""
from __future__ import annotations

import time
from collections.abc import Callable
from datetime import date, timedelta

import requests

from worker.naver_client import API_BASE, TIMEOUT_SECONDS, _check

QNA_URL = f"{API_BASE}/v1/contents/qnas"
INQUIRY_URL = f"{API_BASE}/v1/pay-user/inquiries"

QNA_PAGE_SIZE = 100  # most the Q&A list returns per page
INQUIRY_PAGE_SIZE = 200  # most the inquiry list returns per page
INQUIRY_MAX_DAYS = 365  # "최대 365일만 조회 가능" — the span is counted inclusive of both ends
MAX_PAGES = 200  # guard against runaway paging; a year of Q&A is about 11 pages
RATE_LIMIT_STATUS = 429  # GW.RATE_LIMIT: too many calls in a short time
RETRY_WAITS = (1.0, 2.0, 4.0, 8.0)  # seconds to wait before each retry after a rate-limit refusal


def _get(http: requests.Session, url: str, headers: dict, params: dict,
         sleep: Callable[[float], None] = time.sleep) -> dict:
    """GET with a few spaced retries when the API says we are calling too fast."""
    resp = http.get(url, headers=headers, params=params, timeout=TIMEOUT_SECONDS)
    for wait in RETRY_WAITS:
        if resp.status_code != RATE_LIMIT_STATUS:
            break
        sleep(wait)
        resp = http.get(url, headers=headers, params=params, timeout=TIMEOUT_SECONDS)
    return _check(resp)


def _spans(start: date, end: date, max_days: int) -> list[tuple[date, date]]:
    """Split [start, end] (both inclusive) into spans of at most max_days days."""
    spans = []
    cursor = start
    while cursor <= end:
        span_end = min(cursor + timedelta(days=max_days - 1), end)
        spans.append((cursor, span_end))
        cursor = span_end + timedelta(days=1)
    return spans


def fetch_qnas(token: str, start: date, end: date, session: requests.Session | None = None,
               sleep: Callable[[float], None] = time.sleep) -> list[dict]:
    """Product Q&A posted between start and end (inclusive); answered ones carry `answer`."""
    http = session or requests.Session()
    headers = {"Authorization": f"Bearer {token}"}
    base = {"fromDate": f"{start:%Y-%m-%d}T00:00:00.000+09:00",
            "toDate": f"{end:%Y-%m-%d}T23:59:59.000+09:00", "size": QNA_PAGE_SIZE}
    items: list[dict] = []
    for page in range(1, MAX_PAGES + 1):
        data = _get(http, QNA_URL, headers, {**base, "page": page}, sleep)
        items += data.get("contents") or []
        if data.get("last", True) or page >= (data.get("totalPages") or 0):
            break
    return list({str(i.get("questionId")): i for i in items}.values())


def fetch_customer_inquiries(token: str, start: date, end: date,
                             session: requests.Session | None = None,
                             sleep: Callable[[float], None] = time.sleep) -> list[dict]:
    """1:1 customer inquiries registered between start and end (inclusive), split into 365-day calls."""
    http = session or requests.Session()
    headers = {"Authorization": f"Bearer {token}"}
    items: list[dict] = []
    for span_start, span_end in _spans(start, end, INQUIRY_MAX_DAYS):
        base = {"startSearchDate": f"{span_start:%Y-%m-%d}", "endSearchDate": f"{span_end:%Y-%m-%d}",
                "size": INQUIRY_PAGE_SIZE}
        for page in range(1, MAX_PAGES + 1):
            data = _get(http, INQUIRY_URL, headers, {**base, "page": page}, sleep)
            items += data.get("content") or []
            if data.get("last", True):
                break
    return list({str(i.get("inquiryNo")): i for i in items}.values())
