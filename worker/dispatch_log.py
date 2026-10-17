"""Shared work log of 송장 registrations: Telegram-started runs are written to dispatch/queue.json too.

log_run() never raises and never delays the registration itself: when the data repository cannot be
reached the entry waits in a local outbox file, and flush_outbox() (run by the bot's relay loop) sends
it later. store.dispatch_queue.record() is idempotent, so a retry never duplicates an entry.
"""
from __future__ import annotations

import json
import logging
from collections.abc import Mapping, Sequence
from datetime import datetime
from pathlib import Path

from store import dispatch_queue
from store.base import DataStore
from worker.dispatch_client import DispatchResult
from worker.paths import DATA_DIR

OUTBOX_FILE = DATA_DIR / "발주기록" / "작업로그_미전송.json"

log = logging.getLogger("telegram_bot")


def result_dict(result: DispatchResult, now: datetime) -> dict:
    """The queue's `result` shape for a finished dispatch."""
    return {"processed_at": now.isoformat(timespec="seconds"), "simulated": result.simulated,
            "success_ids": list(result.success_ids),
            "failures": [{"id": f.id, "code": f.code, "message": f.message} for f in result.failures],
            "error": None}


def failed_dict(error: str, now: datetime) -> dict:
    """The queue's `result` shape for a dispatch that raised."""
    return {"processed_at": now.isoformat(timespec="seconds"), "simulated": False,
            "success_ids": [], "failures": [], "error": error}


def _read_outbox(path: Path) -> dict[str, dict]:
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


def _write_outbox(entries: Mapping[str, dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(dict(entries), ensure_ascii=False, indent=1), encoding="utf-8")


def _send(store: DataStore, request_id: str, entry: Mapping) -> None:
    dispatch_queue.record(store, request_id, entry["items"], entry["by"],
                          datetime.fromisoformat(entry["created_at"]), entry["result"], entry["status"],
                          entry.get("source", dispatch_queue.SOURCE_TELEGRAM))


def _keep(request_id: str, entry: Mapping, outbox: Path) -> None:
    try:
        _write_outbox({**_read_outbox(outbox), request_id: dict(entry)}, outbox)
    except Exception as exc:  # noqa: BLE001  # the log is optional work: the registration already happened
        log.error("작업 기록 임시 저장 실패 %s: %s", request_id, type(exc).__name__)


def log_run(store: DataStore | None, request_id: str, items: Sequence[Mapping], by: str, created_at: datetime,
            result_dict: Mapping, status: str, outbox: Path = OUTBOX_FILE) -> bool:
    """Add a finished run to the shared log; True when it reached the data repository."""
    if store is None:
        return False
    entry = {"items": [dict(i) for i in items], "by": by, "created_at": created_at.isoformat(timespec="seconds"),
             "result": dict(result_dict), "status": status, "source": dispatch_queue.SOURCE_TELEGRAM}
    try:
        _send(store, request_id, entry)
        return True
    except Exception as exc:  # noqa: BLE001  # only the type: a message can hold a token URL
        log.error("작업 기록 저장 실패 %s: %s (나중에 다시 보냅니다)", request_id, type(exc).__name__)
        _keep(request_id, entry, outbox)
        return False


def flush_outbox(store: DataStore, outbox: Path = OUTBOX_FILE) -> int:
    """Send the entries that could not be stored earlier; returns how many went through."""
    try:
        waiting = _read_outbox(outbox)
    except Exception as exc:  # noqa: BLE001  # an unreadable outbox is retried next round
        log.error("미전송 작업 기록을 읽지 못했습니다: %s", type(exc).__name__)
        return 0
    sent: list[str] = []
    for request_id, entry in waiting.items():
        try:
            _send(store, request_id, entry)
            sent.append(request_id)
        except Exception as exc:  # noqa: BLE001  # keep this one, go on with the rest
            log.error("작업 기록 재전송 실패 %s: %s", request_id, type(exc).__name__)
    if sent:
        _write_outbox({k: v for k, v in waiting.items() if k not in sent}, outbox)
    return len(sent)
