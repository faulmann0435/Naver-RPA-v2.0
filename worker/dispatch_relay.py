"""Relay of the app's 송장 등록 requests: the app queues them in the data repository, this PC sends them.

process_queue() runs every ~RELAY_INTERVAL seconds inside the Telegram bot's polling loop. Each pending
request is dispatched at most once: a local record (앱요청_처리기록.json) is written BEFORE anything is
sent and kept even if saving the result to the data repository fails, so a request that still looks
pending there is never sent a second time.
"""
from __future__ import annotations

import json
import logging
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

import tomllib

from store import dispatch_queue as queue
from store.base import DataStore
from store.github_store import GitHubStore
from worker.dispatch_client import DispatchResult
from worker.dispatch_log import OUTBOX_FILE, failed_dict, flush_outbox, result_dict
from worker.invoice_run import dispatch_report, register
from worker.naver_adapter import KST
from worker.order_run import SECRETS_FILE
from worker.paths import DATA_DIR

RELAY_INTERVAL = 30  # seconds between two looks at the queue
PROCESSED_FILE = DATA_DIR / "발주기록" / "앱요청_처리기록.json"
WRITE_FAILED_NOTICE = "⚠️ 앱 요청 결과를 저장하지 못했습니다. GitHub 토큰에 쓰기 권한이 필요합니다. 개발자에게 알려 주세요."
INTERRUPTED = "처리 도중 봇이 멈춰 결과를 알 수 없습니다. 네이버에서 등록 여부를 확인하세요."
HEADER = "🚀 앱에서 요청한 송장 등록 ({by})"
STARTED = "started"

log = logging.getLogger("telegram_bot")


def load_store(secrets_file: Path = SECRETS_FILE) -> DataStore | None:
    """The app's data repository from the secrets file; None when there is no data_store section."""
    if not secrets_file.exists():
        return None
    section = tomllib.loads(secrets_file.read_text(encoding="utf-8")).get("data_store")
    return GitHubStore.from_secrets(section) if section else None


def _read_processed(path: Path) -> dict[str, dict]:
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


def _remember(path: Path, request_id: str, record: dict) -> None:
    records = {**_read_processed(path), request_id: record}
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8")


def _dispatch(request: dict, register_fn: Callable[..., DispatchResult]) -> DispatchResult:
    queue.validate_items(request["items"])
    entry = {"when": str(request["created_at"]), "matched": request["items"], "includable": []}
    return register_fn(entry, False)


def _notify(notify: Callable[[str], None], text: str) -> None:
    try:
        notify(text)
    except Exception:  # a Telegram hiccup must not undo a finished dispatch
        log.exception("앱 요청 알림 실패")


def _save(store: DataStore, request_id: str, record: dict, notify: Callable[[str], None], warn: bool) -> bool:
    try:
        queue.complete(store, request_id, record["result"], record["status"])
        return True
    except Exception as exc:  # noqa: BLE001  # permission, network or conflict: the local record protects us
        log.error("앱 요청 %s 결과 저장 실패: %s", request_id, type(exc).__name__)
        if warn:
            _notify(notify, WRITE_FAILED_NOTICE)
        return False


def _handle(store: DataStore, request: dict, register_fn: Callable[..., DispatchResult],
            notify: Callable[[str], None], now: Callable[[], datetime], path: Path) -> None:
    request_id, report = request["id"], ""
    _remember(path, request_id, {"status": STARTED})  # before sending: a crash must not cause a second send
    try:
        result = _dispatch(request, register_fn)
        record = {"status": queue.STATUS_DONE, "result": result_dict(result, now())}
        report = dispatch_report(result)
    except Exception as exc:  # never retried: sending again could double-register
        log.exception("앱 요청 %s 처리 실패", request_id)
        error = f"{type(exc).__name__}: {str(exc)[:300]}"
        record = {"status": queue.STATUS_FAILED, "result": failed_dict(error, now())}
        report = f"⚠️ 등록 실패: {error}"
    _remember(path, request_id, record)
    saved = _save(store, request_id, record, notify, warn=True)
    _notify(notify, f"{HEADER.format(by=request.get('by', ''))}\n{report}")
    if not saved:
        log.warning("앱 요청 %s: 결과는 텔레그램으로만 알렸습니다", request_id)


def _retry_save(store: DataStore, request_id: str, record: dict, notify: Callable[[str], None],
                now: Callable[[], datetime], path: Path) -> None:
    """A request processed here but still pending in the queue: only try to store its result again."""
    if record.get("status") == STARTED:
        record = {"status": queue.STATUS_FAILED, "result": failed_dict(INTERRUPTED, now())}
        _remember(path, request_id, record)
    _save(store, request_id, record, notify, warn=False)


def process_queue(store: DataStore, register_fn: Callable[..., DispatchResult] = register,
                  notify: Callable[[str], None] = lambda text: None,
                  now: Callable[[], datetime] = lambda: datetime.now(KST),
                  processed_file: Path = PROCESSED_FILE) -> int:
    """Process every pending request, oldest first; returns how many were dispatched in this call."""
    count = 0
    for request in queue.pending(store):
        known = _read_processed(processed_file).get(request["id"])
        if known is not None:
            _retry_save(store, request["id"], known, notify, now, processed_file)
            continue
        _handle(store, request, register_fn, notify, now, processed_file)
        count += 1
    return count


def make_relay(store: DataStore, notify: Callable[[str], None], outbox: Path = OUTBOX_FILE) -> Callable[[], None]:
    """The zero-argument callable poll_forever runs every RELAY_INTERVAL seconds (queue + unsent work log)."""
    def relay() -> None:
        process_queue(store, notify=notify)
        try:
            flush_outbox(store, outbox)
        except Exception as exc:  # noqa: BLE001  # the work log is optional: never stop the bot for it
            log.error("미전송 작업 기록 처리 실패: %s", type(exc).__name__)

    return relay
