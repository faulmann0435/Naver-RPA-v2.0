"""Telegram remote control: [📦 발주 처리] sends the purchase orders to the owner's phone.

    python -m worker.telegram_bot

.env keys: TELEGRAM_BOT_TOKEN, TELEGRAM_ALLOWED_CHAT_IDS (comma separated). Until a chat id is
allowed, the bot only answers with that chat's id so it can be added. Only allowed chats can
run anything. Nothing is sent to the delivery sites: the owner shares the files on KakaoTalk
and taps [✅ 카톡 공유 완료], which records the orders so they are not sent twice.
[📋 발송 안 된 주문 전부] ignores that record and puts every unshipped order back (30 days).

송장: when the owner forwards a tracking file (한진 file or a supplier's list, xlsx/xls), the bot matches it to the unshipped Naver orders,
sends the filled upload xlsx and a summary with buttons. Tapping [🚚 송장 등록] registers the tracking
numbers on Naver (worker.dispatch_client; really sent only when INVOICE_DISPATCH_ENABLED=1).
"""
from __future__ import annotations

import hashlib
import io
import logging
import logging.handlers
import os
import sys
import time
import uuid
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

import requests

from worker.dispatch_client import DispatchResult
from worker.dispatch_relay import RELAY_INTERVAL, load_store, make_relay
from worker.invoice_run import (
    PENDING_FILE as INVOICE_PENDING_FILE,
)
from worker.invoice_run import (
    InvoiceRun,
    add_pending,
    dispatch_report,
    invoice_summary,
    prepare_invoices,
    register,
    remove_pending,
)
from worker.invoice_run import read_pending as read_invoice_pending
from worker.naver_client import ENV_FILE, read_env_file
from worker import single_instance
from worker.order_run import (
    PENDING_FILE,
    RunResult,
    mark_shared,
    read_pending,
    run_all_unshipped,
    run_orders,
    summary,
    write_pending,
)
from worker.paths import DATA_DIR

API = "https://api.telegram.org/bot{token}/{method}"
RUN_BUTTON = "📦 발주 처리"
ALL_BUTTON = "📋 발송 안 된 주문 전부"
DONE_PREFIX = "done:"
PENDING_SENT = "sent:"  # + the asked-about runs, so an old question cannot settle a newer run
PENDING_REDO = "redo:"
INVOICE_PREFIX = "inv:"  # inv:<run id>:m (matched only) | a (matched + checkable) | x (cancel)
FILE_API = "https://api.telegram.org/file/bot{token}/{path}"
POLL_TIMEOUT = 30
HTTP_TIMEOUT = POLL_TIMEOUT + 10
TOKEN_KEY = "TELEGRAM_BOT_TOKEN"
CHATS_KEY = "TELEGRAM_ALLOWED_CHAT_IDS"
BACKOFF_START, BACKOFF_MAX = 5, 60  # seconds to wait after a failed getUpdates, doubling up to the max
LOCK_NAME = "telegram_bot"
LOG_FILE_NAME = "telegram_bot.log"
LOG_MAX_BYTES, LOG_BACKUPS = 1_000_000, 5
IP_NOT_ALLOWED = "IP_NOT_ALLOWED"  # Naver gateway: this PC's public IP is not registered
CONFLICT_MARK = "Conflict"  # Telegram: another getUpdates is running (a second copy of the bot)

log = logging.getLogger("telegram_bot")


def _setting(key: str) -> str:
    return os.environ.get(key) or read_env_file(ENV_FILE).get(key, "")


@dataclass(frozen=True)
class BotSettings:
    token: str
    allowed_chats: frozenset[int]

    @classmethod
    def load(cls) -> BotSettings:
        token = _setting(TOKEN_KEY)
        if not token:
            raise SystemExit(f".env에 {TOKEN_KEY}가 없습니다.")
        chats = frozenset(int(c) for c in _setting(CHATS_KEY).replace(" ", "").split(",") if c)
        return cls(token, chats)


class Telegram:
    """The few Bot API calls the bot needs."""

    def __init__(self, token: str, session: requests.Session | None = None):
        self._token = token
        self._http = session or requests.Session()

    def call(self, method: str, files: dict | None = None, **params) -> dict:
        url = API.format(token=self._token, method=method)
        resp = self._http.post(url, data=params if files else None, json=None if files else params,
                               files=files, timeout=HTTP_TIMEOUT)
        body = resp.json()
        if not body.get("ok"):
            raise RuntimeError(f"Telegram {method}: {body.get('description')}")
        return body["result"]

    def send(self, chat_id: int, text: str, **extra) -> dict:
        return self.call("sendMessage", chat_id=chat_id, text=text, **extra)

    def send_file(self, chat_id: int, name: str, data: bytes) -> dict:
        return self.call("sendDocument", files={"document": (name, data)}, chat_id=chat_id)

    def download(self, file_id: str) -> bytes:
        """Bytes of a file sent to the bot. Errors never carry the URL, which contains the token."""
        path = self.call("getFile", file_id=file_id).get("file_path")
        if not path:
            raise RuntimeError("Telegram getFile: file_path 가 없습니다.")
        try:
            resp = self._http.get(FILE_API.format(token=self._token, path=path), timeout=HTTP_TIMEOUT)
        except requests.RequestException as exc:
            raise RuntimeError(f"Telegram 파일 내려받기 실패: {type(exc).__name__}") from None
        if not resp.ok:
            raise RuntimeError(f"Telegram 파일 내려받기 실패: HTTP {resp.status_code}")
        return resp.content


class TelegramNotifier:
    """Notifier for the downloads watcher: one message per line to every allowed chat."""

    def __init__(self, settings: BotSettings):
        self._tg = Telegram(settings.token)
        self._chats = settings.allowed_chats

    def send(self, text: str) -> None:
        for chat in self._chats:
            self._tg.send(chat, text)


def failure_message(exc: Exception, what: str = "발주 처리") -> str:
    """What the owner sees when a run fails; known causes get a plain-language hint."""
    text = f"⚠️ {what} 실패: {type(exc).__name__}: {str(exc)[:300]}"
    if IP_NOT_ALLOWED in str(exc):
        text += ("\n\n이 PC의 인터넷 주소(IP)가 바뀌어 네이버가 막았습니다. "
                 "커머스API센터에서 이 PC의 새 IP를 등록해야 합니다. 개발자에게 알려 주세요.")
    return text


def run_keyboard() -> dict:
    return {"keyboard": [[{"text": RUN_BUTTON}, {"text": ALL_BUTTON}]], "resize_keyboard": True,
            "is_persistent": True}


def pending_token(pending: dict[str, dict]) -> str:
    """Short fingerprint of the waiting runs; fits Telegram's 64-byte callback data."""
    return hashlib.sha1(",".join(sorted(pending)).encode()).hexdigest()[:12]


class OrderBot:
    def __init__(self, tg: Telegram, allowed: frozenset[int],
                 runner: Callable[[], RunResult] = run_orders,
                 all_runner: Callable[[], RunResult] = run_all_unshipped,
                 mark: Callable[[tuple[str, ...]], int] = mark_shared,
                 pending_file: Path = PENDING_FILE,
                 invoice_preparer: Callable[[bytes], InvoiceRun] = prepare_invoices,
                 invoice_register: Callable[..., DispatchResult] = register,
                 invoice_pending_file: Path = INVOICE_PENDING_FILE):
        self.tg, self.allowed, self.runner, self.mark = tg, allowed, runner, mark
        self.all_runner = all_runner
        self.pending_file = pending_file  # runs awaiting [공유 완료]; a file so a restart keeps them
        self.invoice_preparer, self.invoice_register = invoice_preparer, invoice_register
        self.invoice_pending_file = invoice_pending_file  # 송장 runs awaiting a button tap

    @property
    def pending(self) -> dict[str, dict]:
        return read_pending(self.pending_file)

    def _save_pending(self, pending: dict[str, dict]) -> None:
        write_pending(pending, self.pending_file)

    def handle(self, update: dict) -> None:
        if "callback_query" in update:
            self._on_button(update["callback_query"])
        elif "message" in update:
            self._on_message(update["message"])

    def _on_message(self, message: dict) -> None:
        chat = message["chat"]["id"]
        if chat not in self.allowed:
            log.info("등록 안 된 대화방 %s", chat)
            self.tg.send(chat, f"이 대화방 번호는 {chat} 입니다. 개발자에게 알려 주세요.")
            return
        if message.get("document"):
            self._on_document(chat, message["document"])
        elif message.get("text") == RUN_BUTTON:
            if self.pending:
                self._ask_about_pending(chat)
            else:
                self._run(chat)
        elif message.get("text") == ALL_BUTTON:
            self._run(chat, self.all_runner)  # deliberately ignores the ledger, so no pending question
        else:
            self.tg.send(chat, f"아래 [{RUN_BUTTON}] 버튼을 누르면 새 주문으로 발주서를 만듭니다.\n"
                               f"[{ALL_BUTTON}]은 전에 공유한 주문까지, 송장 안 들어간 주문을 모두 넣습니다.\n"
                               "한진 송장 엑셀 파일을 이 대화방에 보내면 송장을 등록합니다.",
                         reply_markup=run_keyboard())

    def _on_document(self, chat: int, document: dict) -> None:
        if not str(document.get("file_name", "")).lower().endswith((".xlsx", ".xls")):
            self.tg.send(chat, "송장 엑셀 파일(.xlsx 또는 .xls)을 보내 주세요.")
            return
        self.tg.send(chat, "⏳ 송장 파일 확인 중…")
        try:
            run = self.invoice_preparer(self.tg.download(document["file_id"]))
        except Exception as exc:  # report every failure to the owner instead of going silent
            log.exception("송장 확인 실패")
            self.tg.send(chat, failure_message(exc, "송장 확인"))
            return
        for f in run.files:
            self.tg.send_file(chat, f.filename, f.data)
        self.tg.send(chat, invoice_summary(run), **self._invoice_markup(run))

    def _invoice_markup(self, run: InvoiceRun) -> dict:
        matched, checks = len(run.result.matched()), len(run.result.includable_checks())
        if not matched and not checks:
            return {"reply_markup": run_keyboard()}
        run_id = add_pending(run, self.invoice_pending_file)
        buttons = []
        if matched:
            buttons.append((f"🚚 일치 {matched}건 송장 등록", "m"))
        if checks:
            buttons.append((f"🚚 확인 필요 {checks}건 포함 ({matched + checks}건) 등록", "a"))
        buttons.append(("✖ 취소", "x"))
        rows = [[{"text": text, "callback_data": f"{INVOICE_PREFIX}{run_id}:{code}"}] for text, code in buttons]
        return {"reply_markup": {"inline_keyboard": rows}}

    def _run(self, chat: int, runner: Callable[[], RunResult] | None = None) -> None:
        self.tg.send(chat, "⏳ 네이버에서 주문을 가져오는 중…")
        try:
            run = (runner or self.runner)()
        except Exception as exc:  # report every failure to the owner instead of going silent
            log.exception("발주 처리 실패")
            self.tg.send(chat, failure_message(exc))
            return
        for f in run.files:
            self.tg.send_file(chat, f["filename"], f["data"].getvalue())
        extra: dict = {"reply_markup": run_keyboard()}
        if run.product_order_ids:
            run_id = uuid.uuid4().hex[:12]
            entry = {"ids": list(run.product_order_ids), "when": f"{run.when:%m/%d %H:%M}"}
            self._save_pending({**self.pending, run_id: entry})
            button = {"text": "✅ 카톡 공유 완료", "callback_data": f"{DONE_PREFIX}{run_id}"}
            extra = {"reply_markup": {"inline_keyboard": [[button]]}}
        self.tg.send(chat, summary(run), **extra)

    def _ask_about_pending(self, chat: int) -> None:
        """A run is still unconfirmed: settle it first, or its orders would land on two purchase orders."""
        waiting = self.pending
        count = len({i for entry in waiting.values() for i in entry["ids"]})
        whens = ", ".join(entry["when"] for entry in waiting.values())
        token = pending_token(waiting)
        buttons = [[{"text": "✅ 보냈음 → 기록하고 새 주문만", "callback_data": f"{PENDING_SENT}{token}"}],
                   [{"text": "↩️ 안 보냈음 → 그 주문까지 다시 만들기", "callback_data": f"{PENDING_REDO}{token}"}]]
        self.tg.send(chat, f"⚠️ 지난 발주서({whens}, 주문 {count}건)가 아직 [카톡 공유 완료] 처리되지 않았습니다.\n"
                           "그 발주서를 카톡으로 보냈나요?", reply_markup={"inline_keyboard": buttons})

    def _on_button(self, query: dict) -> None:
        chat = query["message"]["chat"]["id"]
        data = query.get("data", "")
        if chat not in self.allowed:
            return
        if data.startswith((PENDING_SENT, PENDING_REDO)):
            self.tg.call("answerCallbackQuery", callback_query_id=query["id"])
            token = data.removeprefix(PENDING_SENT).removeprefix(PENDING_REDO)
            self._settle_pending(chat, query, token, sent=data.startswith(PENDING_SENT))
        elif data.startswith(INVOICE_PREFIX):
            self.tg.call("answerCallbackQuery", callback_query_id=query["id"])
            self._on_invoice_button(chat, query, data.removeprefix(INVOICE_PREFIX))
        elif data.startswith(DONE_PREFIX):
            self.tg.call("answerCallbackQuery", callback_query_id=query["id"])
            self._confirm_run(chat, query, data.removeprefix(DONE_PREFIX))

    def _clear_buttons(self, chat: int, query: dict) -> None:
        try:
            self.tg.call("editMessageReplyMarkup", chat_id=chat, message_id=query["message"]["message_id"],
                         reply_markup={"inline_keyboard": []})
        except RuntimeError:  # already cleared (a double tap): the reply below matters more
            log.info("버튼 지우기 생략: message %s", query["message"]["message_id"])

    def _on_invoice_button(self, chat: int, query: dict, data: str) -> None:
        run_id, _, choice = data.partition(":")
        self._clear_buttons(chat, query)
        entry = read_invoice_pending(self.invoice_pending_file).get(run_id)
        if entry is None or choice not in ("m", "a", "x"):
            self.tg.send(chat, "이 버튼은 만료됐습니다. 필요하면 송장 파일을 다시 보내 주세요.")
            return
        try:
            if choice == "x":
                self.tg.send(chat, "취소했습니다. 등록하지 않았습니다.", reply_markup=run_keyboard())
                return
            self.tg.send(chat, "⏳ 네이버에 송장을 등록하는 중…")
            result = self.invoice_register(entry, include_checks=(choice == "a"))
            self.tg.send(chat, dispatch_report(result), reply_markup=run_keyboard())
        except Exception as exc:  # report every failure to the owner instead of going silent
            log.exception("송장 등록 실패")
            self.tg.send(chat, failure_message(exc, "송장 등록"))
        finally:
            remove_pending(run_id, self.invoice_pending_file)

    def _confirm_run(self, chat: int, query: dict, run_id: str) -> None:
        waiting = self.pending
        if run_id not in waiting:
            self.tg.send(chat, "이 버튼은 만료됐습니다. 필요하면 다시 [발주 처리]를 눌러 주세요.")
            return
        added = self.mark(tuple(waiting[run_id]["ids"]))
        self._save_pending({k: v for k, v in waiting.items() if k != run_id})
        self._clear_buttons(chat, query)
        self.tg.send(chat, f"✅ 주문 {added}건을 공유 완료로 기록했습니다. 다음 [발주 처리]에서는 빠집니다.",
                     reply_markup=run_keyboard())

    def _settle_pending(self, chat: int, query: dict, token: str, sent: bool) -> None:
        waiting = self.pending
        self._clear_buttons(chat, query)
        if not waiting or pending_token(waiting) != token:
            self.tg.send(chat, "이미 처리된 질문입니다.", reply_markup=run_keyboard())
            return
        if sent:
            added = self.mark(tuple(sorted({i for entry in waiting.values() for i in entry["ids"]})))
            self.tg.send(chat, f"✅ 지난 주문 {added}건을 공유 완료로 기록했습니다. 새 주문만 가져옵니다.")
        else:
            self.tg.send(chat, "↩️ 지난 발주서 파일은 버려 주세요. 그 주문까지 포함해 새로 만듭니다.")
        self._save_pending({})
        self._run(chat)


def _run_relay(relay: Callable[[], None]) -> None:
    try:
        relay()
    except Exception:  # the app's requests are optional work: log it and keep the bot alive
        log.exception("앱 송장 요청 처리 실패")


def poll_forever(bot: OrderBot, sleep: Callable[[float], None] = time.sleep,
                 relay: Callable[[], None] | None = None,
                 clock: Callable[[], float] = time.monotonic) -> None:
    """Long-poll Telegram forever; every RELAY_INTERVAL seconds also run relay (the app's 송장 requests)."""
    offset = 0
    delay = 0
    conflict_logged = False
    last_relay: float | None = None
    while True:
        if relay is not None and (last_relay is None or clock() - last_relay >= RELAY_INTERVAL):
            last_relay = clock()
            _run_relay(relay)
        try:
            updates = bot.tg.call("getUpdates", offset=offset, timeout=POLL_TIMEOUT)
        except (requests.RequestException, RuntimeError) as exc:
            delay = BACKOFF_START if delay == 0 else min(delay * 2, BACKOFF_MAX)
            if CONFLICT_MARK in str(exc):
                if not conflict_logged:
                    log.warning("다른 곳에서 같은 봇이 실행 중입니다. 다른 PC나 창을 끄세요. (%s)", exc)
                    conflict_logged = True
            elif isinstance(exc, (requests.Timeout, requests.ConnectionError)):
                # everyday network blips: one line, and only the type (the message can hold the token URL)
                log.warning("텔레그램 연결 지연, %d초 뒤 다시 시도 (%s)", delay, type(exc).__name__)
            else:
                log.exception("텔레그램 연결 실패, 다시 시도")
            sleep(delay)
            continue
        delay, conflict_logged = 0, False
        for update in updates:
            offset = update["update_id"] + 1
            try:
                bot.handle(update)
            except Exception:  # one bad update must not stop the bot
                log.exception("처리 실패: update %s", update.get("update_id"))


def setup_logging(log_dir: Path | None) -> None:
    """Log to the console when there is one (pythonw has none) and, given a folder, to a rotating file."""
    handlers: list[logging.Handler] = []
    if sys.stderr is not None:
        handlers.append(logging.StreamHandler())
    if log_dir is not None:
        log_dir.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.handlers.RotatingFileHandler(
            log_dir / LOG_FILE_NAME, maxBytes=LOG_MAX_BYTES, backupCount=LOG_BACKUPS, encoding="utf-8"))
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s", handlers=handlers)


def _build_relay(settings: BotSettings) -> Callable[[], None] | None:
    """Relay for the app's 송장 requests, or None without a data_store section (or when it is unusable)."""
    try:
        store = load_store()
    except Exception as exc:  # a broken secrets file must not keep the bot from starting
        log.error("앱 요청 연결 안 함: 데이터 저장소 설정을 읽지 못했습니다 (%s)", type(exc).__name__)
        return None
    if store is None:
        log.info("데이터 저장소 설정이 없어 앱의 송장 요청은 받지 않습니다.")
        return None
    log.info("앱의 송장 요청을 %d초마다 확인합니다.", RELAY_INTERVAL)
    return make_relay(store, TelegramNotifier(settings).send)


def main() -> int:
    if isinstance(sys.stdout, io.TextIOWrapper):  # None under pythonw
        sys.stdout.reconfigure(encoding="utf-8")
    record_dir = DATA_DIR / "발주기록"
    try:
        lock = single_instance.acquire(LOCK_NAME, record_dir)
    except single_instance.AlreadyRunning:
        setup_logging(None)  # leave the running copy's log file alone
        log.error("발주봇이 이미 실행 중입니다. 새로 시작하지 않습니다.")
        return 1
    with lock:
        setup_logging(record_dir / "로그")
        try:
            settings = BotSettings.load()
            bot = OrderBot(Telegram(settings.token), settings.allowed_chats)
            me = bot.tg.call("getMe")
            log.info("봇 @%s 시작 (허용된 대화방 %d개). 끄려면 Ctrl+C", me["username"], len(settings.allowed_chats))
            poll_forever(bot, relay=_build_relay(settings))
        except BaseException as exc:  # under pythonw the log file is the only place a crash shows up
            if not isinstance(exc, KeyboardInterrupt):
                log.exception("봇이 멈췄습니다 (작업 스케줄러가 1분 뒤 다시 시작)")
            raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
