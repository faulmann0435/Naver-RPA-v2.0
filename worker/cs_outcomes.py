"""How each customer question ended, and the 📊 CS 성적 numbers made from that.

Outcomes: ai_as_is (AI draft posted unchanged), owner_edited (the AI draft posted after the owner
changed it), owner_written (the owner's own text posted where there was no AI draft),
seller_center (the owner chose to answer there), answered_elsewhere (answered in the seller center
without any choice in Telegram). Every outcome goes to outcomes.jsonl in the CS state dir (local,
unmasked) and, masked, to cs/outcomes.json in the data repository. Like worker.dispatch_log, the
store write never raises: on a failure the outcome waits in a local outbox and flush_outbox() sends
it later (idempotent by id).
"""
from __future__ import annotations

import json
import logging
from collections.abc import Mapping
from datetime import datetime, timedelta
from pathlib import Path

from store.base import Author, ConflictError, DataStore
from worker.cs_drafter import ROUTE_AI, ROUTE_HUMAN, mask_question
from worker.cs_state import OUTCOMES_FILE, append_jsonl, read_jsonl, short_id

AI_AS_IS, OWNER_EDITED, OWNER_WRITTEN = "ai_as_is", "owner_edited", "owner_written"
SELLER_CENTER, ANSWERED_ELSEWHERE = "seller_center", "answered_elsewhere"
OUTCOMES_PATH = "cs/outcomes.json"
OUTBOX_FILE = "outcomes_미전송.json"
MAX_STORED = 1000
MAX_ATTEMPTS = 3
AUTHOR = Author("PC 봇", "app@naver-rpa.local")
RECENT_DAYS = 7
HINT = "AI 그대로 비율이 충분히 높아지면 자동 등록을 검토할 수 있습니다."

log = logging.getLogger("telegram_bot")


def build_outcome(key: str, entry: Mapping, outcome: str, now: datetime, *, simulated: bool = False,
                  owner_answer: str = "") -> dict:
    """One outcome record. No names or order ids exist in `entry`, so none can end up here."""
    draft = entry.get("draft") or {}
    return {"id": short_id(key), "channel": entry.get("channel", ""),
            "outcome": outcome, "route": draft.get("route"), "kind": draft.get("kind"),
            "simulated": simulated, "created": entry.get("created", ""),
            "settled_at": now.isoformat(timespec="seconds"), "product_name": entry.get("product_name", ""),
            "question": entry.get("question", ""), "ai_draft": draft.get("draft"), "owner_answer": owner_answer}


def mask_outcome(outcome: Mapping) -> dict:
    """Copy for the data repository: phone numbers and long digit runs hidden in all free text."""
    masked = dict(outcome)
    for field in ("question", "ai_draft", "owner_answer"):
        if masked.get(field):
            masked[field] = mask_question(masked[field])
    return masked


# --- data repository + outbox -----------------------------------------------------------------

def _read_outbox(path: Path) -> list[dict]:
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else []


def _write_outbox(path: Path, outcomes: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(outcomes, ensure_ascii=False, indent=1), encoding="utf-8")


def _store_append(store: DataStore, outcomes: list[dict]) -> None:
    """Add outcomes (skipping ids already there) to cs/outcomes.json; retries on a write conflict."""
    for attempt in range(MAX_ATTEMPTS):
        snapshot = store.read_text(OUTCOMES_PATH)
        current = json.loads(snapshot.content) if snapshot else []
        known = {o.get("id") for o in current}
        fresh = [o for o in outcomes if o["id"] not in known]
        if not fresh:
            return
        content = json.dumps((current + fresh)[-MAX_STORED:], ensure_ascii=False, indent=1)
        try:
            store.write_text(OUTCOMES_PATH, content, snapshot.sha if snapshot else None,
                             "CS 문의 결과 기록", AUTHOR)
            return
        except ConflictError:
            if attempt == MAX_ATTEMPTS - 1:
                raise


def record_outcome(state_dir: Path, store: DataStore | None, outcome: Mapping) -> bool:
    """Write locally, then to the data repository; True when the repository got it. Never raises."""
    try:
        append_jsonl(state_dir, OUTCOMES_FILE, dict(outcome))
    except Exception as exc:  # noqa: BLE001  # the owner's flow must go on
        log.error("CS 결과 로컬 기록 실패: %s", type(exc).__name__)
    if store is None:
        return False
    masked = mask_outcome(outcome)
    try:
        _store_append(store, [masked])
        return True
    except Exception as exc:  # noqa: BLE001  # only the type: a message can hold a token URL
        log.error("CS 결과 저장 실패: %s (나중에 다시 보냅니다)", type(exc).__name__)
        try:
            _write_outbox(state_dir / OUTBOX_FILE, _read_outbox(state_dir / OUTBOX_FILE) + [masked])
        except Exception as inner:  # noqa: BLE001
            log.error("CS 결과 임시 저장 실패: %s", type(inner).__name__)
        return False


def flush_outbox(state_dir: Path, store: DataStore) -> int:
    """Send outcomes that could not be stored earlier; returns how many went through. Never raises."""
    path = state_dir / OUTBOX_FILE
    try:
        waiting = _read_outbox(path)
        if not waiting:
            return 0
        _store_append(store, waiting)
        path.unlink()
        return len(waiting)
    except Exception as exc:  # noqa: BLE001  # retried next round
        log.error("미전송 CS 결과 처리 실패: %s", type(exc).__name__)
        return 0


# --- 📊 CS 성적 -------------------------------------------------------------------------------

def _pct(part: int, whole: int) -> str:
    return f"{round(part * 100 / whole)}%" if whole else "-"


def _choices(outcomes: list[dict]) -> list[dict]:
    """Questions where the owner judged an AI draft in Telegram (answered_elsewhere had no choice,
    and without a draft there was nothing to judge)."""
    return [o for o in outcomes if o["outcome"] != ANSWERED_ELSEWHERE and (o.get("ai_draft") or "").strip()]


def _count(outcomes: list[dict], kind: str) -> int:
    return sum(1 for o in outcomes if o["outcome"] == kind)


def _section(title: str, outcomes: list[dict]) -> str:
    if not outcomes:
        return f"{title}: 아직 없습니다."
    choices = _choices(outcomes)
    n = len(choices)
    lines = [f"{title}: 총 {len(outcomes)}건",
             f"✅ AI 그대로 {_count(choices, AI_AS_IS)}건 ({_pct(_count(choices, AI_AS_IS), n)})"
             f" · ✏️ 고쳐서 올림 {_count(choices, OWNER_EDITED)}건 ({_pct(_count(choices, OWNER_EDITED), n)})"
             f" · ✍️ 직접 씀 {_count(choices, OWNER_WRITTEN)}건 ({_pct(_count(choices, OWNER_WRITTEN), n)})"
             f" · 🙋 판매자센터 {_count(choices, SELLER_CENTER)}건 ({_pct(_count(choices, SELLER_CENTER), n)})"]
    elsewhere = _count(outcomes, ANSWERED_ELSEWHERE)
    if elsewhere:
        lines.append(f"(판매자센터에서 따로 답함 {elsewhere}건은 비율에서 뺐습니다)")
    no_draft = sum(1 for o in outcomes if o["outcome"] != ANSWERED_ELSEWHERE and o not in choices)
    if no_draft:
        lines.append(f"(AI 초안이 없던 {no_draft}건은 비율에서 뺐습니다)")
    for route in (ROUTE_AI, ROUTE_HUMAN):
        group = [o for o in choices if o.get("route") == route]
        if group:
            lines.append(f"· {route} 분류 {len(group)}건 중 ✅ {_count(group, AI_AS_IS)}건 "
                         f"({_pct(_count(group, AI_AS_IS), len(group))})")
    return "\n".join(lines)


def stats_text(state_dir: Path, now: datetime) -> str:
    """The 📊 CS 성적 message: last 7 days and all time."""
    everything = read_jsonl(state_dir, OUTCOMES_FILE)
    cutoff = now - timedelta(days=RECENT_DAYS)
    recent = [o for o in everything if datetime.fromisoformat(o["settled_at"]) >= cutoff]
    parts = ["📊 CS 성적", _section("최근 7일", recent), _section("전체", everything)]
    if any(o.get("simulated") for o in everything):
        parts.append("(시험 모드로 올린 건도 포함됩니다)")
    parts.append(HINT)
    return "\n\n".join(parts)
