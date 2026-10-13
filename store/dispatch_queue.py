"""Queue of 송장 등록 requests between the app (which cannot call Naver) and the owner's PC.

The app appends a request to dispatch/queue.json in the data repository; the bot on the owner's PC
processes pending requests and writes the result back. Pure functions over any DataStore.

    {"requests": [{"id", "created_at", "by", "status": "pending"|"done"|"failed",
                   "items": [{"ids": [...], "tracking", "courier", "recipient"}],
                   "result": null | {"processed_at", "simulated", "success_ids", "failures", "error"}}]}
"""
from __future__ import annotations

import json
import re
import uuid
from collections.abc import Callable, Mapping, Sequence
from datetime import datetime

from core.invoice_match import COURIER_CODES
from store.base import Author, ConflictError, DataStore

DISPATCH_QUEUE_PATH = "dispatch/queue.json"
STATUS_PENDING, STATUS_DONE, STATUS_FAILED = "pending", "done", "failed"
MAX_SHIPMENTS = 500
MAX_REQUESTS = 100
MAX_ATTEMPTS = 3
TRACKING_PATTERN = re.compile(r"[A-Za-z0-9-]{1,50}")
AUTHOR_EMAIL = "app@naver-rpa.local"


def validate_items(items: Sequence[Mapping]) -> None:
    """Refuse a malformed shipment list. Raises ValueError (Korean)."""
    if not items:
        raise ValueError("등록할 송장이 없습니다.")
    if len(items) > MAX_SHIPMENTS:
        raise ValueError(f"한 번에 {MAX_SHIPMENTS}건까지만 보낼 수 있습니다. ({len(items)}건)")
    for item in items:
        ids = item.get("ids") or []
        if not ids or not all(isinstance(i, str) and i.isascii() and i.isdigit() for i in ids):
            raise ValueError(f"상품주문번호가 숫자가 아닙니다: {ids}")
        tracking = item.get("tracking")
        if not isinstance(tracking, str) or not TRACKING_PATTERN.fullmatch(tracking):
            raise ValueError(f"송장번호 형식이 올바르지 않습니다: {tracking!r}")
        if item.get("courier") not in set(COURIER_CODES.values()):
            raise ValueError(f"알 수 없는 택배사 코드입니다: {item.get('courier')!r}")


def _pairs(items: Sequence[Mapping]) -> frozenset[tuple[str, str]]:
    return frozenset((i, str(item["tracking"])) for item in items for i in item["ids"])


def _load(store: DataStore) -> tuple[list[dict], str | None]:
    snapshot = store.read_text(DISPATCH_QUEUE_PATH)
    if snapshot is None:
        return [], None
    try:
        requests = json.loads(snapshot.content).get("requests", [])
    except (ValueError, AttributeError):
        raise ValueError("dispatch/queue.json 형식이 올바르지 않습니다.") from None
    return list(requests), snapshot.sha


def _trim(requests: list[dict]) -> list[dict]:
    """Keep the newest MAX_REQUESTS; only finished requests are ever dropped."""
    excess = len(requests) - MAX_REQUESTS
    kept: list[dict] = []
    for request in requests:
        if excess > 0 and request["status"] != STATUS_PENDING:
            excess -= 1
            continue
        kept.append(request)
    return kept


def _save(store: DataStore, requests: list[dict], sha: str | None, message: str, by: str) -> None:
    content = json.dumps({"requests": _trim(requests)}, ensure_ascii=False, indent=1)
    store.write_text(DISPATCH_QUEUE_PATH, content, sha, message, Author(by, AUTHOR_EMAIL))


def _update(store: DataStore, change: Callable[[list[dict]], str | None], message: str, by: str) -> str | None:
    """Read, apply change(requests) (returns a value to stop early), write; retry on a write conflict."""
    for attempt in range(MAX_ATTEMPTS):
        requests, sha = _load(store)
        outcome = change(requests)
        if outcome is not None:
            return outcome
        try:
            _save(store, requests, sha, message, by)
            return None
        except ConflictError:
            if attempt == MAX_ATTEMPTS - 1:
                raise
    return None  # pragma: no cover


def _blocks_resubmit(request: dict) -> bool:
    """A pending request, or one really done without failures, must not be sent twice."""
    if request["status"] == STATUS_PENDING:
        return True
    result = request.get("result") or {}
    return request["status"] == STATUS_DONE and not result.get("simulated") and not result.get("failures")


def submit(store: DataStore, items: Sequence[Mapping], by: str, now: datetime) -> str:
    """Append a pending request and return its id (the existing id for an exact duplicate)."""
    clean = [{"ids": list(i["ids"]), "tracking": str(i["tracking"]), "courier": i["courier"],
              "recipient": str(i.get("recipient", ""))} for i in items]
    validate_items(clean)
    wanted = _pairs(clean)
    new_id = uuid.uuid4().hex[:12]

    def add(requests: list[dict]) -> str | None:
        for request in requests:
            if _blocks_resubmit(request) and _pairs(request["items"]) == wanted:
                return str(request["id"])
        requests.append({"id": new_id, "created_at": now.isoformat(timespec="seconds"), "by": by,
                         "status": STATUS_PENDING, "items": clean, "result": None})
        return None

    return _update(store, add, f"송장 등록 요청 {new_id}", by) or new_id


def pending(store: DataStore) -> list[dict]:
    """Pending requests, oldest first."""
    return [r for r in _load(store)[0] if r["status"] == STATUS_PENDING]


def get(store: DataStore, request_id: str) -> dict | None:
    return next((r for r in _load(store)[0] if r["id"] == request_id), None)


def complete(store: DataStore, request_id: str, result: Mapping, status: str) -> None:
    """Store the result of a request and set its status (done or failed)."""
    if status not in (STATUS_DONE, STATUS_FAILED):
        raise ValueError(f"잘못된 상태입니다: {status}")
    found: list[dict] = []

    def apply(requests: list[dict]) -> str | None:
        found.clear()
        for index, request in enumerate(requests):
            if request["id"] == request_id:
                requests[index] = {**request, "status": status, "result": dict(result)}
                found.append(request)
        return None if found else "missing"

    if _update(store, apply, f"송장 등록 결과 {request_id}", "PC 봇") == "missing":
        raise KeyError(f"요청을 찾을 수 없습니다: {request_id}")
