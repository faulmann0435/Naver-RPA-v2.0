"""The owner's side of a CS alert in Telegram: buttons, typed answers, posting, outcome records.

Used by worker.telegram_bot.OrderBot (same bot process, same token). An alert (see worker.cs_watch)
has buttons; this module handles what the owner taps or types:

  ✅ 이대로 올리기  -> "이 답변을 올릴까요?" + [올리기] [취소]
  ✏️ 초안 고쳐서 올리기 -> sends the AI draft as a tap-to-copy block, then waits for the next text of
                      that chat (or a Telegram reply to the alert) -> "이렇게 올릴까요?"
                      + [올리기] [다시 고치기] [취소]; 다시 고치기 sends the owner's last version
  ✍️ 직접 쓰기     -> the same without a block to copy, for a question without an AI draft
  🙋 판매자센터에서 -> the owner answers in the seller center; buttons removed

Waiting/preview state is kept in cs_pending.json (a restart keeps it, 24 hours). A preview is
one-shot: it is removed before posting, so a second tap on [올리기] finds nothing. Whether the answer
really reaches Naver is decided by CS_ANSWER_ENABLED (see worker.cs_answer).
"""
from __future__ import annotations

import hashlib
import html
import logging
from collections.abc import Callable, Mapping
from datetime import datetime
from pathlib import Path
from typing import Any

from store.base import DataStore
from worker.cs_alert import (
    ACT_AI,
    ACT_CANCEL,
    ACT_POST,
    ACT_REWRITE,
    ACT_SELLER,
    ACT_WRITE,
    CS_PREFIX,
    SAFE_LIMIT,
    preview_keyboard,
)
from worker.cs_answer import (
    AlreadyAnswered,
    AnswerResult,
    CsAnswerError,
    submit_answer,
    validate_answer,
)
from worker.cs_outcomes import (
    AI_AS_IS,
    ANSWERED_ELSEWHERE,
    OWNER_EDITED,
    OWNER_WRITTEN,
    SELLER_CENTER,
    build_outcome,
    record_outcome,
    stats_text,
)
from worker.cs_state import (
    find_by_message,
    find_by_sid,
    load_state,
    read_pending,
    save_state,
    write_pending,
)
from worker.naver_adapter import KST

STATS_BUTTON, STATS_COMMAND, CANCEL_COMMAND = "📊 CS 성적", "/성적", "/취소"
SIMULATED_NOTICE = ("🧪 시험 모드: 네이버에 올리지 않았습니다.\n"
                    "(실제로 올리려면 개발자가 켜야 합니다. 이 문의는 판매자센터에서 직접 답변해 주세요.)")
POSTED_NOTICE = "✅ 네이버에 답변을 올렸습니다."
EXPIRED = "이 버튼은 만료됐습니다."
SETTLED = "이미 처리된 문의입니다."
ASK_TEXT = "답변을 입력해 주세요 (취소하려면 /취소)"
ASK_EDIT = "✏️ 아래 글을 눌러 복사한 뒤, 고쳐서 보내 주세요 (취소하려면 /취소)"
SOURCE_AI, SOURCE_OWNER = "ai", "owner"

log = logging.getLogger("telegram_bot")


def default_answer(entry: Mapping, text: str) -> AnswerResult:
    return submit_answer(entry["channel"], entry["id"], entry.get("created", ""), text)


def fingerprint(text: str) -> str:
    return hashlib.sha1(text.encode()).hexdigest()[:6]


def _draft_text(entry: Mapping) -> str:
    return ((entry.get("draft") or {}).get("draft") or "").strip()


def _outcome_of(posted: str, draft: str) -> str:
    """The draft unchanged is AI 그대로 (also when the owner sent it back as is); a changed text is an
    edit of the draft when there was one, else the owner's own answer."""
    if draft and posted == draft:
        return AI_AS_IS
    return OWNER_EDITED if draft else OWNER_WRITTEN


class CsFlow:
    def __init__(self, tg: Any, state_dir: Path, store: DataStore | None = None,
                 answer_fn: Callable[[Mapping, str], AnswerResult] = default_answer,
                 now: Callable[[], datetime] = lambda: datetime.now(KST)):
        self.tg, self.state_dir, self.store = tg, state_dir, store
        self.answer_fn, self.now = answer_fn, now

    # --- state helpers --------------------------------------------------------------------

    def _state(self) -> dict[str, dict]:
        return load_state(self.state_dir)[0]

    def _pending(self) -> dict[str, dict]:
        return read_pending(self.state_dir, self.now())

    def _set_waiting(self, chat: int, sid: str | None) -> None:
        pending = self._pending()
        waiting = {k: v for k, v in pending["waiting"].items() if k != str(chat)}
        if sid:
            waiting[str(chat)] = {"sid": sid, "since": self.now().isoformat()}
        write_pending(self.state_dir, {**pending, "waiting": waiting})

    def _set_preview(self, chat: int, sid: str, preview: dict | None) -> None:
        pending = self._pending()
        previews = {k: v for k, v in pending["previews"].items() if k != f"{chat}:{sid}"}
        if preview:
            previews[f"{chat}:{sid}"] = {**preview, "since": self.now().isoformat()}
        write_pending(self.state_dir, {**pending, "previews": previews})

    def _send(self, chat: int, text: str, **extra: Any) -> None:
        try:
            self.tg.send(chat, text, **extra)
        except Exception as exc:  # noqa: BLE001  # a Telegram hiccup must not break the flow
            log.error("CS 안내 전송 실패: %s", type(exc).__name__)

    def _clear_alert_buttons(self, entry: Mapping) -> None:
        for chat, message_id in (entry.get("messages") or {}).items():
            try:
                self.tg.call("editMessageReplyMarkup", chat_id=int(chat), message_id=message_id,
                             reply_markup={"inline_keyboard": []})
            except Exception:  # noqa: BLE001  # already cleared, or the message is gone
                log.info("CS 알림 버튼 지우기 생략: message %s", message_id)

    def _settle(self, key: str, outcome: str, *, simulated: bool = False, owner_answer: str = "") -> None:
        """Record the outcome in the state entry, the local log and the data repository."""
        state, _ = load_state(self.state_dir)
        entry = state[key]
        now = self.now()
        save_state(self.state_dir, {**state, key: {**entry, "outcome": outcome, "outcome_simulated": simulated,
                                                   "outcome_at": now.isoformat(timespec="seconds")}})
        record_outcome(self.state_dir, self.store, build_outcome(key, entry, outcome, now, simulated=simulated,
                                                                 owner_answer=owner_answer))
        self._clear_alert_buttons(entry)

    # --- buttons --------------------------------------------------------------------------

    def on_button(self, chat: int, query: dict, data: str) -> None:
        parts = data.removeprefix(CS_PREFIX).split(":")
        action, sid = parts[0], (parts[1] if len(parts) > 1 else "")
        found = find_by_sid(self._state(), sid)
        if found is None:
            self._send(chat, EXPIRED)
            return
        key, entry = found
        if entry.get("outcome"):
            self._send(chat, SETTLED)
            self._clear_alert_buttons(entry)
            return
        handlers = {ACT_AI: self._ask_ai, ACT_WRITE: self._ask_write, ACT_REWRITE: self._ask_write,
                    ACT_SELLER: self._seller_center, ACT_CANCEL: self._cancel}
        if action == ACT_POST:
            self._post(chat, key, entry, parts[2] if len(parts) > 2 else "")
        elif action in handlers:
            handlers[action](chat, key, entry)

    def _ask_ai(self, chat: int, key: str, entry: dict) -> None:
        text = _draft_text(entry)
        if not text:
            self._send(chat, "올릴 AI 초안이 없습니다. [✍️ 직접 쓰기]를 눌러 주세요.")
            return
        self._show_preview(chat, entry, text, SOURCE_AI)

    def _show_preview(self, chat: int, entry: Mapping, text: str, source: str) -> bool:
        """Ask "올릴까요?"; False (with a Korean reason sent) when the text cannot be posted at all."""
        try:
            validate_answer(text)
        except CsAnswerError as err:
            self._send(chat, f"⚠️ {err}")
            return False
        sid = entry["sid"]
        self._set_preview(chat, sid, {"text": text, "source": source})
        title = "이 답변을 올릴까요?" if source == SOURCE_AI else "이렇게 올릴까요?"
        self._send(chat, f"{title}\n\n{text}",
                   reply_markup=preview_keyboard(sid, fingerprint(text), rewrite=source == SOURCE_OWNER))
        return True

    def _ask_write(self, chat: int, key: str, entry: dict) -> None:
        """Wait for the owner's text; first send the text to start from (their last version, else the
        AI draft) as a <pre> block, which Telegram copies with one tap."""
        sid = entry["sid"]
        preview = self._pending()["previews"].get(f"{chat}:{sid}") or {}
        base = (preview.get("text") or _draft_text(entry)).strip()
        self._set_preview(chat, sid, None)
        self._set_waiting(chat, sid)
        if not base:
            self._send(chat, ASK_TEXT)
            return
        self._send(chat, ASK_EDIT)
        self._send(chat, f"<pre>{html.escape(base[:SAFE_LIMIT], quote=False)}</pre>", parse_mode="HTML")

    def _cancel(self, chat: int, key: str, entry: dict) -> None:
        self._set_preview(chat, entry["sid"], None)
        waiting = self._pending()["waiting"].get(str(chat))
        if waiting and waiting.get("sid") == entry["sid"]:
            self._set_waiting(chat, None)
        self._send(chat, "취소했습니다. 알림의 버튼은 그대로입니다.")

    def _seller_center(self, chat: int, key: str, entry: dict) -> None:
        self._settle(key, SELLER_CENTER)
        self._send(chat, "🙋 알겠습니다. 판매자센터에서 직접 답변해 주세요.")

    def _post(self, chat: int, key: str, entry: dict, fp: str) -> None:
        sid = entry["sid"]
        preview = self._pending()["previews"].get(f"{chat}:{sid}")
        if preview is None or fingerprint(preview["text"]) != fp:
            self._send(chat, EXPIRED)
            return
        self._set_preview(chat, sid, None)  # one-shot: a second tap finds nothing
        text = preview["text"].strip()
        try:
            result = self.answer_fn(entry, text)
        except AlreadyAnswered as err:
            self._settle(key, ANSWERED_ELSEWHERE)
            self._send(chat, f"⚠️ {err}")
            return
        except CsAnswerError as err:
            self._send(chat, f"⚠️ 답변을 올리지 못했습니다. {err}")
            self._show_preview(chat, entry, text, preview["source"])  # lets the owner retry
            return
        except Exception as exc:  # noqa: BLE001  # report every failure to the owner instead of going silent
            log.exception("CS 답변 등록 실패")
            self._send(chat, f"⚠️ 답변을 올리지 못했습니다: {type(exc).__name__}: {str(exc)[:200]}")
            self._show_preview(chat, entry, text, preview["source"])
            return
        outcome = _outcome_of(text, _draft_text(entry))
        self._settle(key, outcome, simulated=result.simulated, owner_answer="" if outcome == AI_AS_IS else text)
        self._send(chat, SIMULATED_NOTICE if result.simulated else POSTED_NOTICE)

    # --- typed text -----------------------------------------------------------------------

    def _target(self, chat: int, message: dict) -> tuple[str, dict] | None:
        """The question a text of this chat answers: the replied-to alert first, else the one waited for."""
        state = self._state()
        reply_to = (message.get("reply_to_message") or {}).get("message_id")
        if reply_to is not None and (hit := find_by_message(state, chat, reply_to)):
            return hit
        waiting = self._pending()["waiting"].get(str(chat))
        return find_by_sid(state, waiting["sid"]) if waiting else None

    def handles_text(self, chat: int, message: dict) -> bool:
        """True when this text belongs to the CS flow (so the bot's normal reply must not run)."""
        text = (message.get("text") or "").strip()
        if not text or (text.startswith("/") and text != CANCEL_COMMAND):
            return False
        if text == CANCEL_COMMAND:
            return str(chat) in self._pending()["waiting"]
        return self._target(chat, message) is not None

    def on_text(self, chat: int, message: dict) -> None:
        text = (message.get("text") or "").strip()
        if text == CANCEL_COMMAND:
            self._set_waiting(chat, None)
            self._send(chat, "취소했습니다.")
            return
        target = self._target(chat, message)
        if target is None:
            return
        key, entry = target
        if entry.get("outcome"):
            self._set_waiting(chat, None)
            self._send(chat, SETTLED)
            return
        if self._show_preview(chat, entry, text, SOURCE_OWNER):
            self._set_waiting(chat, None)
        else:  # too long or empty: keep waiting for a better text
            self._set_waiting(chat, entry["sid"])

    # --- stats ----------------------------------------------------------------------------

    def stats_text(self) -> str:
        return stats_text(self.state_dir, self.now())
