"""Read-only Naver Commerce API client: access token, recently paid orders, order details.

Nothing here changes an order on the store (no confirm, no dispatch).
Keys come from the environment or a .env file next to the repo root (git-ignored):
    NAVER_COMMERCE_CLIENT_ID, NAVER_COMMERCE_CLIENT_SECRET
"""
from __future__ import annotations

import base64
import os
import time
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

import bcrypt
import requests

from worker.naver_adapter import KST

API_BASE = "https://api.commerce.naver.com/external"
TOKEN_URL = f"{API_BASE}/v1/oauth2/token"
ORDERS_BASE = f"{API_BASE}/v1/pay-order/seller/product-orders"
LAST_CHANGED_URL = f"{ORDERS_BASE}/last-changed-statuses"
ORDER_QUERY_URL = f"{ORDERS_BASE}/query"

TIMEOUT_SECONDS = 15
QUERY_CHUNK = 300  # most product orders one query call takes
WINDOW = timedelta(hours=24)  # widest lastChanged range the API accepts
MAX_PAGES = 30  # guard against runaway paging
ENV_FILE = Path(__file__).resolve().parent.parent / ".env"


class NaverApiError(RuntimeError):
    """The API refused a call; the message carries Naver's code and text."""


@dataclass(frozen=True)
class Credentials:
    client_id: str
    client_secret: str


def read_env_file(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    pairs = (line.split("=", 1) for line in path.read_text(encoding="utf-8").splitlines() if "=" in line)
    return {k.strip(): v.strip() for k, v in pairs if not k.strip().startswith("#")}


def load_credentials(env_file: Path = ENV_FILE) -> Credentials:
    """Keys from the environment first, then the .env file; fail clearly when missing."""
    file_values = read_env_file(env_file)
    pick = lambda key: os.environ.get(key) or file_values.get(key, "")
    creds = Credentials(pick("NAVER_COMMERCE_CLIENT_ID"), pick("NAVER_COMMERCE_CLIENT_SECRET"))
    if not creds.client_id or not creds.client_secret:
        raise NaverApiError("NAVER_COMMERCE_CLIENT_ID / NAVER_COMMERCE_CLIENT_SECRET 가 설정되지 않았습니다 (.env 확인).")
    return creds


def make_signature(creds: Credentials, timestamp_ms: int) -> str:
    """base64(bcrypt(client_id_timestamp, salt=client_secret)) — Naver's electronic signature."""
    password = f"{creds.client_id}_{timestamp_ms}".encode()
    return base64.b64encode(bcrypt.hashpw(password, creds.client_secret.encode())).decode()


def _check(resp: requests.Response) -> dict:
    if resp.ok:
        return resp.json()
    try:
        body = resp.json()
        detail = f"{body.get('code', '')} {body.get('message', '')}".strip()
    except ValueError:
        detail = resp.text[:300]
    raise NaverApiError(f"HTTP {resp.status_code}: {detail}")


def fetch_token(creds: Credentials, session: requests.Session | None = None) -> str:
    """Access token for the seller's own store (type SELF)."""
    http = session or requests.Session()
    timestamp_ms = int(time.time() * 1000)
    data = {
        "client_id": creds.client_id,
        "timestamp": timestamp_ms,
        "grant_type": "client_credentials",
        "client_secret_sign": make_signature(creds, timestamp_ms),
        "type": "SELF",
    }
    token = _check(http.post(TOKEN_URL, data=data, timeout=TIMEOUT_SECONDS)).get("access_token")
    if not token:
        raise NaverApiError("토큰 응답에 access_token 이 없습니다.")
    return token


def _api_time(moment: datetime) -> str:
    return moment.astimezone(KST).strftime("%Y-%m-%dT%H:%M:%S.000+09:00")


def _windows(start: datetime, end: datetime) -> list[tuple[datetime, datetime]]:
    edges = []
    cursor = start
    while cursor < end:
        edges.append((cursor, min(cursor + WINDOW, end)))
        cursor += WINDOW
    return edges


def fetch_changed_ids(token: str, start: datetime, end: datetime, changed_type: str = "PAYED",
                      session: requests.Session | None = None) -> list[str]:
    """Product order ids whose status changed to `changed_type` between start and end (oldest first)."""
    http = session or requests.Session()
    headers = {"Authorization": f"Bearer {token}"}
    ids: list[str] = []
    pages = 0
    for w_start, w_end in _windows(start, end):
        params = {"lastChangedFrom": _api_time(w_start), "lastChangedTo": _api_time(w_end),
                  "lastChangedType": changed_type, "limitCount": QUERY_CHUNK}
        while pages < MAX_PAGES:
            pages += 1
            data = _check(http.get(LAST_CHANGED_URL, headers=headers, params=params,
                                   timeout=TIMEOUT_SECONDS)).get("data") or {}
            ids += [str(s["productOrderId"]) for s in data.get("lastChangeStatuses") or [] if s.get("productOrderId")]
            more = data.get("more") or {}
            if not more.get("moreFrom"):
                break
            params = {**params, "lastChangedFrom": more["moreFrom"], "moreSequence": more.get("moreSequence")}
    return list(dict.fromkeys(ids))


def fetch_details(token: str, product_order_ids: Iterable[str],
                  session: requests.Session | None = None) -> list[dict]:
    """Order details (the API `data` items) for the given product order ids, 300 per call."""
    http = session or requests.Session()
    headers = {"Authorization": f"Bearer {token}"}
    ids = [str(i) for i in product_order_ids]
    items: list[dict] = []
    for offset in range(0, len(ids), QUERY_CHUNK):
        body = {"productOrderIds": ids[offset:offset + QUERY_CHUNK]}
        items += _check(http.post(ORDER_QUERY_URL, headers=headers, json=body,
                                  timeout=TIMEOUT_SECONDS)).get("data") or []
    return items
