"""Watch for new unanswered customer questions and alert the owner on Telegram (read-only).

    python -m worker.cs_watch [--once] [--every 10] [--days 2] [--dry-run]

Every cycle: new unanswered product Q&A and 1:1 inquiries are linked to their orders, an AI draft
is written from the store document, and one Telegram message per question goes to the allowed
chats. Nothing is ever posted to Naver. When a question we alerted on is later answered by the
owner, the pair (AI draft, owner answer) is appended to pairs.jsonl for later scoring.

The same cycle also runs inside the Telegram bot (make_bot_job, every 10 minutes when
CS_WATCH_ENABLED=1); there the alerts carry buttons that worker.cs_flow handles. This command line
version sends plain alerts without buttons.

State lives outside the repo (customer data): CS자료/알림기록/ next to the repo, or .env CS_STATE_DIR.
  notified.json  {key: {sid, channel, id, created, product_name, question, draft, notified_at, alerted,
                 order_kind, messages, done, outcome?}}
  pairs.jsonl    one line per alerted question that was answered afterwards
  outcomes.jsonl how each question ended (see worker.cs_outcomes)
"""
from __future__ import annotations

import argparse
import io
import logging
import os
import sys
import time
from collections.abc import Callable
from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import requests

from store.base import DataStore
from worker.cs_alert import NO_AI_NOTE, alert_keyboard, format_alert, trim
from worker.cs_drafter import (
    DraftError,
    DraftResult,
    InquiryView,
    load_store_doc,
    make_client,
    model_name,
)
from worker.cs_drafter import draft as make_draft
from worker.cs_outcomes import (
    AI_AS_IS,
    ANSWERED_ELSEWHERE,
    OWNER_WRITTEN,
    build_outcome,
    flush_outbox,
    record_outcome,
)
from worker.cs_state import (
    CHANNEL_INQUIRY,
    CHANNEL_QNA,
    append_pair,
    load_state,
    save_state,
    short_id,
)
from worker.inquiry_client import fetch_customer_inquiries, fetch_qnas
from worker.naver_adapter import KST
from worker.naver_client import (
    ENV_FILE,
    fetch_details,
    fetch_token,
    load_credentials,
    read_env_file,
)
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
from worker.telegram_bot import BotSettings, Telegram

CS_DIR_NAME, STATE_DIR_NAME = "CS자료", "알림기록"
STATE_DIR_KEY = "CS_STATE_DIR"
BASELINE_AGE = timedelta(hours=6)  # first run: older unanswered questions are recorded, not alerted
REASON_MAX = 60
PAIR_WINDOW_DAYS = 30  # stop waiting for an owner answer after this many days

log = logging.getLogger("cs_watch")


class CsWatchError(RuntimeError):
    """A cycle cannot run (e.g. no allowed Telegram chat); the message is Korean."""


@dataclass(frozen=True)
class Options:
    state_dir: Path
    days: int = 2
    dry_run: bool = False
    buttons: bool = False  # inline buttons under each alert (only useful when the bot handles them)


@dataclass(frozen=True)
class CycleStats:
    alerted: int = 0
    skipped: int = 0
    paired: int = 0
    failed: int = 0


@dataclass(frozen=True)
class Item:
    """One customer question, from either channel, in one shape."""
    key: str
    channel: str
    id: str
    created: datetime
    created_text: str
    answered: bool
    owner_answer: str
    product_name: str
    question: str
    raw: dict


# --- state ------------------------------------------------------------------------------------

def default_state_dir() -> Path:
    override = os.environ.get(STATE_DIR_KEY) or read_env_file(ENV_FILE).get(STATE_DIR_KEY, "")
    return Path(override) if override else DATA_DIR / CS_DIR_NAME / STATE_DIR_NAME


# --- items ------------------------------------------------------------------------------------

def _answer_text(value: Any) -> str:
    if isinstance(value, dict):
        value = value.get("content") or value.get("answer") or ""
    return str(value or "")


def _item(channel: str, prefix: str, item_id: str, created_text: str | None, raw: dict, *,
          question: str, owner_answer: Any) -> Item:
    created = parse_time(created_text) or datetime.now(KST)
    return Item(key=f"{prefix}:{item_id}", channel=channel, id=item_id, created=created,
                created_text=created_text or "", answered=bool(raw.get("answered")),
                owner_answer=_answer_text(owner_answer), product_name=str(raw.get("productName") or ""),
                question=question, raw=raw)


def inquiry_question(inquiry: dict) -> str:
    """A 1:1 inquiry's title and body together (the title often carries the actual question)."""
    return "\n".join(p for p in (inquiry.get("title") or "", inquiry.get("inquiryContent") or "") if p.strip())


def collect_items(qnas: list[dict], inquiries: list[dict]) -> list[Item]:
    """Both channels as Items, oldest first."""
    items = [_item(CHANNEL_QNA, "qna", str(q.get("questionId")), q.get("createDate"), q,
                   question=q.get("question") or "", owner_answer=q.get("answer")) for q in qnas]
    items += [_item(CHANNEL_INQUIRY, "inq", str(i.get("inquiryNo")), i.get("inquiryRegistrationDateTime"), i,
                    question=inquiry_question(i), owner_answer=i.get("answerContent"))
              for i in inquiries]
    return sorted(items, key=lambda i: i.created)


def link_item(item: Item, details: list[dict]) -> LinkResult:
    if item.channel == CHANNEL_QNA:
        return link_qna(item.raw, details)
    return link_customer_inquiry(item.raw, details)


def _entry(item: Item, now: datetime, *, draft: DraftResult | None = None, alerted: bool = False,
           order_kind: str = "", messages: dict[str, int] | None = None) -> dict:
    return {"sid": short_id(item.key), "channel": item.channel, "id": item.id, "created": item.created_text,
            "product_name": item.product_name, "question": item.question,
            "draft": asdict(draft) if draft else None, "notified_at": now.isoformat(),
            "alerted": alerted, "order_kind": order_kind, "messages": messages or {}, "done": False}


def _pair(item: Item, entry: dict, now: datetime) -> dict:
    draft = entry.get("draft") or {}
    return {"key": item.key, "channel": item.channel, "created": item.created_text,
            "question": item.question, "order_kind": entry.get("order_kind", ""),
            "route": draft.get("route"), "kind": draft.get("kind"), "draft": draft.get("draft"),
            "owner_answer": item.owner_answer, "answered_seen_at": now.isoformat()}


def _posted_by_bot(entry: dict) -> bool:
    """The answer on Naver is the one the bot posted (a pair would only repeat it)."""
    return entry.get("outcome") in (AI_AS_IS, OWNER_WRITTEN) and not entry.get("outcome_simulated")


def _settle_elsewhere(state_dir: Path, store: DataStore | None, item: Item, entry: dict, now: datetime) -> dict:
    """An alerted question answered in the seller center without a choice in Telegram."""
    record_outcome(state_dir, store, build_outcome(item.key, entry, ANSWERED_ELSEWHERE, now,
                                                   owner_answer=item.owner_answer))
    return {**entry, "outcome": ANSWERED_ELSEWHERE, "outcome_simulated": False,
            "outcome_at": now.isoformat(timespec="seconds")}


def record_pairs(state_dir: Path, state: dict[str, dict], answered: list[Item], now: datetime,
                 store: DataStore | None = None) -> tuple[dict[str, dict], int]:
    """Write (AI draft, owner answer) pairs for alerted questions now answered; each only once.
    Alerted questions nobody settled in Telegram also get an answered_elsewhere outcome."""
    paired = 0
    for item in answered:
        entry = state[item.key]
        if entry.get("alerted") and not entry.get("outcome"):
            entry = _settle_elsewhere(state_dir, store, item, entry, now)
        if entry.get("alerted") and entry.get("draft") and not _posted_by_bot(entry):
            append_pair(state_dir, _pair(item, entry, now))
            paired += 1
        state = {**state, item.key: {**entry, "done": True}}
        save_state(state_dir, state)
    return state, paired


# --- one cycle --------------------------------------------------------------------------------

def _describe(channel: str, created: datetime) -> str:
    return f"[{channel}] {created.astimezone(KST):%Y-%m-%d %H:%M}"


def _ai_draft(client: Any, doc: str, model: str, view: InquiryView,
              off_note: str = NO_AI_NOTE) -> tuple[DraftResult | None, str]:
    """(draft, note): the note replaces the draft section of the alert when there is no draft."""
    if client is None:
        return None, off_note
    try:
        return make_draft(client, doc, view, model), NO_AI_NOTE
    except DraftError as err:
        log.warning("AI 초안 실패 %s: %s", _describe(view.channel, view.created), err)
        return None, f"(AI 초안 실패: {trim(str(err), REASON_MAX)})"


def _send_everywhere(tg: Any, chats: frozenset[int], text: str,
                     markup: dict | None = None) -> dict[int, int | None]:
    """{chat: message id} of the chats that got the message; empty when none did (failures are logged)."""
    extra = {"reply_markup": markup} if markup else {}
    sent: dict[int, int | None] = {}
    for chat in sorted(chats):
        try:
            result = tg.call("sendMessage", chat_id=chat, text=text, disable_web_page_preview=True, **extra)
            sent[chat] = result.get("message_id") if isinstance(result, dict) else None
        except (RuntimeError, requests.RequestException, ValueError) as err:
            log.error("텔레그램 전송 실패 (채팅 %s): %s", chat, err)
    return sent


def _fetch_order_details(token: str, pending: list[Item], now: datetime) -> list[dict]:
    details = fetch_recent_details(token, DEFAULT_LOOKBACK_DAYS, now=now)
    inquiries = [i.raw for i in pending if i.channel == CHANNEL_INQUIRY]
    missing = missing_inquiry_ids(inquiries, details)
    if missing:  # 1:1 inquiries name their orders: fetch those directly even when older than the lookback
        details = details + fetch_details(token, missing, RetryingSession())
    return details


def _baseline(state: dict[str, dict], pending: list[Item],
              now: datetime) -> tuple[dict[str, dict], list[Item], list[Item]]:
    """First run: unanswered questions older than 6 hours are recorded as seen without an alert."""
    old = [i for i in pending if now - i.created > BASELINE_AGE]
    state = {**state, **{i.key: _entry(i, now) for i in old}}
    return state, old, [i for i in pending if i not in old]


def fetch_start(state: dict[str, dict], end: date, days: int) -> date:
    """First day to fetch: the last `days` days, stretched back to the oldest alerted question still
    waiting for the owner's answer (so a late answer still gets its pair), at most PAIR_WINDOW_DAYS."""
    start = end - timedelta(days=days - 1)
    waiting = [parse_time(e.get("created")) for e in state.values() if e.get("alerted") and not e.get("done")]
    oldest = min((t.astimezone(KST).date() for t in waiting if t), default=start)
    return max(min(start, oldest), end - timedelta(days=PAIR_WINDOW_DAYS - 1))


def _telegram_target(opts: Options) -> tuple[Any, frozenset[int]]:
    if opts.dry_run:
        return None, frozenset()
    settings = BotSettings.load()
    if not settings.allowed_chats:
        raise CsWatchError("TELEGRAM_ALLOWED_CHAT_IDS가 비어 있어 알림을 보낼 곳이 없습니다 (.env 확인).")
    return Telegram(settings.token), settings.allowed_chats


def run_cycle(opts: Options, client: Any, doc: str, model: str, now: datetime | None = None, *,
              target: tuple[Any, frozenset[int]] | None = None, store: DataStore | None = None,
              ai_off_note: str = NO_AI_NOTE) -> CycleStats:
    """One check. `target` = (Telegram, chats) from the bot; without it the .env bot settings are used."""
    now = now or datetime.now(KST)
    state, first_run = load_state(opts.state_dir)
    token = fetch_token(load_credentials())
    end = now.date()
    start = fetch_start(state, end, opts.days)
    items = collect_items(fetch_qnas(token, start, end), fetch_customer_inquiries(token, start, end))
    answered = [i for i in items if i.answered and i.key in state and not state[i.key].get("done")]
    pending = [i for i in items if not i.answered and i.key not in state]

    paired = 0
    if answered and not opts.dry_run:
        state, paired = record_pairs(opts.state_dir, state, answered, now, store)
    skipped: list[Item] = []
    if first_run:
        state, skipped, pending = _baseline(state, pending, now)
        for item in skipped:
            print(f"처음 실행: 오래된 미답변 문의는 알림 없이 기록만 합니다 {_describe(item.channel, item.created)}")
        if not opts.dry_run:
            save_state(opts.state_dir, state)
    if not pending:
        return CycleStats(skipped=len(skipped), paired=paired)

    tg, chats = target or _telegram_target(opts)
    details = _fetch_order_details(token, pending, now)
    alerted = failed = 0
    for item in pending:
        link = link_item(item, details)
        view = InquiryView(item.channel, item.created, item.product_name, item.question, link)
        result, note = _ai_draft(client, doc, model, view, ai_off_note)
        text = format_alert(view, link, result, note)
        if opts.dry_run:
            print(text + "\n" + "-" * 30)
            alerted += 1
        elif sent := _send_everywhere(tg, chats, text,
                                      alert_keyboard(short_id(item.key), result) if opts.buttons else None):
            messages = {str(chat): mid for chat, mid in sent.items() if mid}
            state = {**state, item.key: _entry(item, now, draft=result, alerted=True, order_kind=link.kind,
                                               messages=messages)}
            save_state(opts.state_dir, state)  # after EACH send: a crash must not re-alert
            alerted += 1
        else:
            failed += 1
    return CycleStats(alerted, len(skipped), paired, failed)


# --- inside the Telegram bot ------------------------------------------------------------------

def make_bot_job(tg: Any, chats: frozenset[int], store: DataStore | None,
                 state_dir: Path | None = None) -> Callable[[], None]:
    """The zero-argument callable the bot runs every CS_INTERVAL seconds: one cycle with buttons."""
    directory = state_dir or default_state_dir()
    try:
        client = make_client()
    except Exception as exc:  # noqa: BLE001  # a broken AI setup must not keep the alerts from going out
        log.error("AI 연결 준비 실패, AI 없이 알립니다: %s", type(exc).__name__)
        client = None

    def job() -> None:
        active, doc, note = client, "", NO_AI_NOTE
        if client is not None:  # re-read each cycle so store-document edits apply without a restart
            try:
                doc = load_store_doc(store=store)
            except DraftError as err:
                log.error("가게 정보 문서를 읽지 못해 AI 없이 알립니다: %s", err)
                active, note = None, f"(AI 초안 꺼짐: {trim(str(err), REASON_MAX)})"
        stats = run_cycle(Options(directory, buttons=True), active, doc, model_name(),
                          target=(tg, chats), store=store, ai_off_note=note)
        if store is not None:
            flush_outbox(directory, store)
        if stats != CycleStats():
            log.info(_summary(stats))

    return job


# --- CLI --------------------------------------------------------------------------------------

def _summary(stats: CycleStats) -> str:
    return (f"알림 {stats.alerted}건, 기록만 {stats.skipped}건, "
            f"답변 쌍 {stats.paired}건, 전송 실패 {stats.failed}건")


def watch_forever(opts: Options, every: int, client: Any, doc: str, model: str,
                  sleep: Callable[[float], None] = time.sleep) -> None:
    """Run a cycle, sleep `every` minutes, repeat; a failed cycle is logged and the loop goes on."""
    while True:
        try:
            if client is not None:  # re-read each cycle so store-document edits apply without a restart
                doc = load_store_doc()
            log.info(_summary(run_cycle(opts, client, doc, model)))
        except Exception as err:  # Naver / Claude / Telegram hiccups must not stop the watcher
            log.error("이번 주기 실패, 다음 주기에 다시 시도합니다: %s", err)
        sleep(every * 60)


def main(argv: list[str] | None = None) -> int:
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--once", action="store_true", help="한 번만 확인하고 끝냅니다 (스케줄러용)")
    parser.add_argument("--every", type=int, default=10, help="확인 주기(분)")
    parser.add_argument("--days", type=int, default=2, help="문의를 볼 최근 일수")
    parser.add_argument("--dry-run", action="store_true", help="전송·기록 없이 알림 문구만 출력")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
    opts = Options(default_state_dir(), max(args.days, 1), args.dry_run)
    try:
        client = make_client()
        doc = load_store_doc() if client else ""
    except DraftError as err:
        print(err)
        return 1
    if client is None:
        print(NO_AI_NOTE)
    model = model_name()
    if not args.once:
        try:
            watch_forever(opts, max(args.every, 1), client, doc, model)
        except KeyboardInterrupt:
            pass
        return 0
    try:
        stats = run_cycle(opts, client, doc, model)
    except Exception as err:
        print(f"확인 실패: {err}")
        return 1
    print(_summary(stats))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
