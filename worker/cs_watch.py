"""Watch for new unanswered customer questions and alert the owner on Telegram (read-only).

    python -m worker.cs_watch [--once] [--every 10] [--days 2] [--dry-run]

Every cycle: new unanswered product Q&A and 1:1 inquiries are linked to their orders, an AI draft
is written from the store document, and one Telegram message per question goes to the allowed
chats. Nothing is ever posted to Naver. When a question we alerted on is later answered by the
owner, the pair (AI draft, owner answer) is appended to pairs.jsonl for later scoring.

State lives outside the repo (customer data): CS자료/알림기록/ next to the repo, or .env CS_STATE_DIR.
  notified.json  {key: {channel, id, created, draft, notified_at, alerted, order_kind, done}}
  pairs.jsonl    one line per alerted question that was answered afterwards
"""
from __future__ import annotations

import argparse
import io
import json
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

from tests.regression_harness import REPO_ROOT
from worker.cs_alert import NO_AI_NOTE, format_alert, trim
from worker.cs_drafter import (
    DraftError,
    DraftResult,
    InquiryView,
    load_store_doc,
    make_client,
    model_name,
)
from worker.cs_drafter import draft as make_draft
from worker.inquiry_client import fetch_customer_inquiries, fetch_qnas
from worker.naver_adapter import KST
from worker.naver_client import ENV_FILE, fetch_details, fetch_token, load_credentials, read_env_file
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
from worker.telegram_bot import BotSettings, Telegram

CS_DIR_NAME, STATE_DIR_NAME = "CS자료", "알림기록"
STATE_DIR_KEY = "CS_STATE_DIR"
NOTIFIED_FILE, PAIRS_FILE = "notified.json", "pairs.jsonl"
CHANNEL_QNA, CHANNEL_INQUIRY = "상품문의", "고객문의"
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
    return Path(override) if override else REPO_ROOT.parent / CS_DIR_NAME / STATE_DIR_NAME


def load_state(state_dir: Path) -> tuple[dict[str, dict], bool]:
    """(state, first_run): first_run is True when notified.json does not exist yet."""
    path = state_dir / NOTIFIED_FILE
    if not path.exists():
        return {}, True
    return json.loads(path.read_text(encoding="utf-8")), False


def save_state(state_dir: Path, state: dict[str, dict]) -> None:
    state_dir.mkdir(parents=True, exist_ok=True)
    tmp = state_dir / (NOTIFIED_FILE + ".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(state_dir / NOTIFIED_FILE)


def append_pair(state_dir: Path, pair: dict) -> None:
    state_dir.mkdir(parents=True, exist_ok=True)
    with (state_dir / PAIRS_FILE).open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(pair, ensure_ascii=False) + "\n")


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
           order_kind: str = "") -> dict:
    return {"channel": item.channel, "id": item.id, "created": item.created_text,
            "draft": asdict(draft) if draft else None, "notified_at": now.isoformat(),
            "alerted": alerted, "order_kind": order_kind, "done": False}


def _pair(item: Item, entry: dict, now: datetime) -> dict:
    draft = entry.get("draft") or {}
    return {"key": item.key, "channel": item.channel, "created": item.created_text,
            "question": item.question, "order_kind": entry.get("order_kind", ""),
            "route": draft.get("route"), "kind": draft.get("kind"), "draft": draft.get("draft"),
            "owner_answer": item.owner_answer, "answered_seen_at": now.isoformat()}


def record_pairs(state_dir: Path, state: dict[str, dict], answered: list[Item],
                 now: datetime) -> tuple[dict[str, dict], int]:
    """Write (AI draft, owner answer) pairs for alerted questions now answered; each only once."""
    paired = 0
    for item in answered:
        entry = state[item.key]
        if entry.get("alerted") and entry.get("draft"):
            append_pair(state_dir, _pair(item, entry, now))
            paired += 1
        state = {**state, item.key: {**entry, "done": True}}
        save_state(state_dir, state)
    return state, paired


# --- one cycle --------------------------------------------------------------------------------

def _describe(channel: str, created: datetime) -> str:
    return f"[{channel}] {created.astimezone(KST):%Y-%m-%d %H:%M}"


def _ai_draft(client: Any, doc: str, model: str, view: InquiryView) -> tuple[DraftResult | None, str]:
    """(draft, note): the note replaces the draft section of the alert when there is no draft."""
    if client is None:
        return None, NO_AI_NOTE
    try:
        return make_draft(client, doc, view, model), NO_AI_NOTE
    except DraftError as err:
        log.warning("AI 초안 실패 %s: %s", _describe(view.channel, view.created), err)
        return None, f"(AI 초안 실패: {trim(str(err), REASON_MAX)})"


def _send_everywhere(tg: Any, chats: frozenset[int], text: str) -> bool:
    """True when at least one chat got the message (failures are logged, not raised)."""
    sent = False
    for chat in sorted(chats):
        try:
            tg.call("sendMessage", chat_id=chat, text=text, disable_web_page_preview=True)
            sent = True
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


def run_cycle(opts: Options, client: Any, doc: str, model: str, now: datetime | None = None) -> CycleStats:
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
        state, paired = record_pairs(opts.state_dir, state, answered, now)
    skipped: list[Item] = []
    if first_run:
        state, skipped, pending = _baseline(state, pending, now)
        for item in skipped:
            print(f"처음 실행: 오래된 미답변 문의는 알림 없이 기록만 합니다 {_describe(item.channel, item.created)}")
        if not opts.dry_run:
            save_state(opts.state_dir, state)
    if not pending:
        return CycleStats(skipped=len(skipped), paired=paired)

    tg, chats = _telegram_target(opts)
    details = _fetch_order_details(token, pending, now)
    alerted = failed = 0
    for item in pending:
        link = link_item(item, details)
        view = InquiryView(item.channel, item.created, item.product_name, item.question, link)
        result, note = _ai_draft(client, doc, model, view)
        text = format_alert(view, link, result, note)
        if opts.dry_run:
            print(text + "\n" + "-" * 30)
            alerted += 1
        elif _send_everywhere(tg, chats, text):
            state = {**state, item.key: _entry(item, now, draft=result, alerted=True, order_kind=link.kind)}
            save_state(opts.state_dir, state)  # after EACH send: a crash must not re-alert
            alerted += 1
        else:
            failed += 1
    return CycleStats(alerted, len(skipped), paired, failed)


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
