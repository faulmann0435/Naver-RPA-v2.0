"""Telegram remote control: who may use it, what a run sends, and the 공유 완료 button."""
import io

import pytest

from worker.order_run import RunResult, mark_shared, read_ledger, run_orders, summary
from worker.telegram_bot import ALL_BUTTON, RUN_BUTTON, OrderBot, run_keyboard

OWNER = 111


class FakeTelegram:
    def __init__(self):
        self.sent = []

    def call(self, method, files=None, **params):
        self.sent.append((method, params))
        return {}

    def send(self, chat_id, text, **extra):
        self.sent.append(("sendMessage", {"chat_id": chat_id, "text": text, **extra}))

    def send_file(self, chat_id, name, data):
        self.sent.append(("sendDocument", {"chat_id": chat_id, "name": name}))

    def texts(self):
        return [p["text"] for m, p in self.sent if m == "sendMessage"]


def _run(ids=("1", "2")):
    files = [{"filename": "속초 발주양식.xlsx", "data": io.BytesIO(b"x"), "vendor": "속초"}]
    return RunResult(files=files, product_order_ids=tuple(ids), paid=len(ids), left_out_claims=1,
                     already_shared=0, unmatched=0, needs_review=0, source="test")


def _message(chat, text):
    return {"message": {"chat": {"id": chat}, "text": text}}


def _answer(tg, n):
    """callback data of the n-th button (0 = 보냈음, 1 = 안 보냈음) on the last question."""
    return tg.sent[-1][1]["reply_markup"]["inline_keyboard"][n][0]["callback_data"]


def _press(data):
    return {"callback_query": {"id": "q", "data": data, "message": {"chat": {"id": OWNER}, "message_id": 5}}}


@pytest.fixture
def pending(tmp_path):
    return tmp_path / "pending.json"


def test_stranger_only_learns_their_chat_id(pending):
    tg, calls = FakeTelegram(), []
    bot = OrderBot(tg, frozenset({OWNER}), runner=lambda: calls.append(1) or _run(), pending_file=pending)
    bot.handle(_message(999, RUN_BUTTON))
    assert calls == [] and "999" in tg.texts()[0]


def test_run_sends_files_then_summary_with_done_button(pending):
    tg = FakeTelegram()
    bot = OrderBot(tg, frozenset({OWNER}), runner=_run, pending_file=pending)
    bot.handle(_message(OWNER, RUN_BUTTON))
    methods = [m for m, _ in tg.sent]
    assert methods == ["sendMessage", "sendDocument", "sendMessage"]
    last = tg.sent[-1][1]
    assert last["text"].startswith("📦 발주서 1개")
    assert last["reply_markup"]["inline_keyboard"][0][0]["callback_data"].startswith("done:")


def test_done_button_records_orders_once(pending):
    tg, marked = FakeTelegram(), []
    bot = OrderBot(tg, frozenset({OWNER}), runner=_run, mark=lambda ids: marked.append(ids) or len(ids),
                  pending_file=pending)
    bot.handle(_message(OWNER, RUN_BUTTON))
    data = tg.sent[-1][1]["reply_markup"]["inline_keyboard"][0][0]["callback_data"]
    bot.handle(_press(data))
    bot.handle(_press(data))
    assert marked == [("1", "2")]
    assert "2건을 공유 완료" in tg.texts()[-2] and "만료" in tg.texts()[-1]


def test_failure_is_reported_not_swallowed(pending):
    tg = FakeTelegram()

    def broken():
        raise RuntimeError("HTTP 403: GW.IP_NOT_ALLOWED")

    OrderBot(tg, frozenset({OWNER}), runner=broken, pending_file=pending).handle(_message(OWNER, RUN_BUTTON))
    assert "IP_NOT_ALLOWED" in tg.texts()[-1]


def test_ledger_keeps_shared_orders_out_of_the_next_run(tmp_path):
    ledger = tmp_path / "ledger.json"
    item = lambda i: {"order": {"orderId": "9"}, "productOrder": {
        "productOrderId": i, "productOrderStatus": "PAYED", "productName": "a", "productOption": "b",
        "quantity": 1, "packageNumber": "7", "productId": "5", "shippingAddress": {"zipCode": "12345", "baseAddress": "서울시 중구"}}}
    engine = lambda: type("E", (), {"config": None, "dictionary": None, "settings": None, "source": "t"})()
    mark_shared(("1",), ledger)
    run = run_orders(fetch=lambda: [item("1")], engine_loader=engine, ledger=ledger)
    assert run.paid == 0 and run.already_shared == 1 and run.files == []
    assert "새로 발주할 주문이 없습니다" in summary(run)
    assert read_ledger(ledger) == {"1"}


def _paid(i, name="홍길동"):
    return {"order": {"orderId": "9"}, "productOrder": {
        "productOrderId": i, "productOrderStatus": "PAYED", "productName": "초벌장어 2kg", "productOption": "b",
        "quantity": 1, "packageNumber": "7", "productId": "5",
        "shippingAddress": {"zipCode": "12345", "name": name, "baseAddress": "서울시 중구"}}}


def _engine():
    return type("E", (), {"config": None, "dictionary": None, "settings": None, "source": "t"})()


def test_summary_names_the_orders_left_out_as_already_shared(tmp_path):
    ledger = tmp_path / "ledger.json"
    mark_shared(("1",), ledger)
    run = run_orders(fetch=lambda: [_paid("1", "김철수")], engine_loader=_engine, ledger=ledger)
    assert run.shared_names == ("김철수(초벌장어 2kg)",)
    assert "이미 공유한 주문 1건 제외: 김철수(초벌장어 2kg)" in summary(run)


def test_include_shared_puts_ledger_orders_back_and_warns(tmp_path, monkeypatch):
    ledger = tmp_path / "ledger.json"
    mark_shared(("1",), ledger)
    monkeypatch.setattr("worker.order_run.process_orders",
                        lambda *a, **k: type("R", (), {"files": [], "stats": {}, "unmatched": []})())
    run = run_orders(fetch=lambda: [_paid("1", "김철수"), _paid("2")], engine_loader=_engine,
                     ledger=ledger, include_shared=True)
    assert run.product_order_ids == ("1", "2") and run.already_shared == 1
    text = summary(run)
    assert "전에 공유한 주문 1건도 다시 넣음: 김철수" in text and "두 번 보내지 않게" in text


def test_summary_cuts_long_name_lists():
    names = tuple(f"고객{i}(상품)" for i in range(13))
    run = RunResult(files=[], product_order_ids=(), paid=0, left_out_claims=0, already_shared=13,
                    unmatched=0, needs_review=0, source="t", shared_names=names)
    assert "외 3건" in summary(run) and "고객12" not in summary(run)


def test_keyboard_has_both_run_buttons():
    assert [b["text"] for b in run_keyboard()["keyboard"][0]] == [RUN_BUTTON, ALL_BUTTON]


def test_all_button_runs_everything_without_the_pending_question(pending):
    tg, marked = FakeTelegram(), []
    everything = _run(("1", "2", "3"))
    bot = OrderBot(tg, frozenset({OWNER}), runner=lambda: _run(), all_runner=lambda: everything,
                   mark=lambda ids: marked.append(ids) or len(ids), pending_file=pending)
    bot.handle(_message(OWNER, RUN_BUTTON))  # leaves an unconfirmed run
    bot.handle(_message(OWNER, ALL_BUTTON))
    assert "아직 [카톡 공유 완료]" not in tg.texts()[-1]
    assert sorted(len(e["ids"]) for e in bot.pending.values()) == [2, 3]


def _bot(tg, pending, runs, marked):
    return OrderBot(tg, frozenset({OWNER}), runner=lambda: runs.pop(0),
                    mark=lambda ids: marked.append(ids) or len(ids), pending_file=pending)


def test_unconfirmed_run_survives_a_restart(pending):
    tg, marked = FakeTelegram(), []
    _bot(tg, pending, [_run()], marked).handle(_message(OWNER, RUN_BUTTON))
    data = tg.sent[-1][1]["reply_markup"]["inline_keyboard"][0][0]["callback_data"]
    _bot(tg, pending, [], marked).handle(_press(data))  # a fresh bot, as after a PC reboot
    assert marked == [("1", "2")]


def test_run_asks_first_when_last_run_is_unconfirmed(pending):
    tg, runs = FakeTelegram(), [_run(), _run(("3",))]
    bot = _bot(tg, pending, runs, [])
    bot.handle(_message(OWNER, RUN_BUTTON))
    bot.handle(_message(OWNER, RUN_BUTTON))
    assert len(runs) == 1  # the second press did not run
    assert "주문 2건" in tg.texts()[-1] and "보냈나요" in tg.texts()[-1]


def test_sent_answer_records_old_orders_then_runs(pending):
    tg, marked = FakeTelegram(), []
    runs = [_run(), _run(("3",))]
    bot = _bot(tg, pending, runs, marked)
    bot.handle(_message(OWNER, RUN_BUTTON))
    bot.handle(_message(OWNER, RUN_BUTTON))
    bot.handle(_press(_answer(tg, 0)))
    assert marked == [("1", "2")] and runs == []
    assert [e["ids"] for e in bot.pending.values()] == [["3"]]


def test_redo_answer_drops_old_run_without_recording(pending):
    tg, marked = FakeTelegram(), []
    runs = [_run(), _run()]
    bot = _bot(tg, pending, runs, marked)
    bot.handle(_message(OWNER, RUN_BUTTON))
    old = tg.sent[-1][1]["reply_markup"]["inline_keyboard"][0][0]["callback_data"]
    bot.handle(_message(OWNER, RUN_BUTTON))
    bot.handle(_press(_answer(tg, 1)))
    assert marked == [] and runs == [] and len(bot.pending) == 1
    bot.handle(_press(old))  # the dropped run's button no longer records anything
    assert marked == [] and "만료" in tg.texts()[-1]


def test_answering_twice_does_not_run_twice(pending):
    tg, runs = FakeTelegram(), [_run(), _run(), _run()]
    bot = _bot(tg, pending, runs, [])
    bot.handle(_message(OWNER, RUN_BUTTON))
    bot.handle(_message(OWNER, RUN_BUTTON))
    sent = _answer(tg, 0)
    bot.handle(_press(sent))
    bot.handle(_press(sent))  # double tap: must not record the new run it just made
    assert len(runs) == 1 and "이미 처리된" in tg.texts()[-1]
    assert [e["ids"] for e in bot.pending.values()] == [["1", "2"]]


# --- poll_forever backoff -----------------------------------------------------------------------

class _Stop(Exception):
    """Raised by the fake to end the endless loop."""


class _ScriptedTelegram:
    def __init__(self, script):
        self._script = list(script)

    def call(self, method, **params):
        step = self._script.pop(0)
        if isinstance(step, Exception):
            raise step
        return step


def _poll(script):
    from worker.telegram_bot import poll_forever

    sleeps: list[float] = []
    bot = OrderBot(_ScriptedTelegram(script), frozenset())
    with pytest.raises(_Stop):
        poll_forever(bot, sleep=sleeps.append)
    return sleeps


def test_poll_backs_off_and_caps():
    fail = RuntimeError("Telegram getUpdates: network down")
    assert _poll([fail] * 7 + [_Stop()]) == [5, 10, 20, 40, 60, 60, 60]


def test_poll_backoff_resets_after_success():
    fail = RuntimeError("down")
    assert _poll([fail, fail, [], fail, _Stop()]) == [5, 10, 5]


def test_poll_conflict_is_warned_once(caplog):
    conflict = RuntimeError("Telegram getUpdates: Conflict: terminated by other getUpdates request")
    with caplog.at_level("WARNING", logger="telegram_bot"):
        sleeps = _poll([conflict] * 3 + [_Stop()])
    assert sleeps == [5, 10, 20]
    assert len([r for r in caplog.records if "실행 중" in r.message]) == 1
    assert all(r.exc_info is None for r in caplog.records)


def test_poll_network_hiccup_is_one_line_without_url(caplog):
    import requests

    hiccup = requests.ConnectionError("HTTPSConnectionPool(host='api.telegram.org'): url: /botSECRET/getUpdates")
    with caplog.at_level("WARNING", logger="telegram_bot"):
        _poll([requests.ReadTimeout("read timed out"), hiccup, _Stop()])
    messages = [r.getMessage() for r in caplog.records]
    assert len(messages) == 2 and all("연결 지연" in m for m in messages)
    assert "5초 뒤" in messages[0] and "10초 뒤" in messages[1]
    assert all(r.exc_info is None for r in caplog.records)
    assert "SECRET" not in caplog.text


def test_failure_message_explains_ip_not_allowed():
    from worker.naver_client import NaverApiError
    from worker.telegram_bot import failure_message

    text = failure_message(NaverApiError("HTTP 403: GW.IP_NOT_ALLOWED 호출이 허용되지 않은 IP입니다."))
    assert "IP_NOT_ALLOWED" in text and "새 IP를 등록" in text
    assert "새 IP" not in failure_message(RuntimeError("boom"))


# --- 송장 (tracking numbers) ----------------------------------------------------------------------

from functools import partial  # noqa: E402

from tests.invoice_helpers import api_item, hanjin_row, hanjin_xlsx  # noqa: E402
from worker.dispatch_client import DispatchFailure, DispatchResult, dispatch_or_simulate  # noqa: E402
from worker.invoice_run import prepare_invoices, register  # noqa: E402
from worker.invoice_run import read_ledger as read_invoice_ledger  # noqa: E402
from worker.invoice_run import read_pending as read_invoice_pending  # noqa: E402


class InvoiceTelegram(FakeTelegram):
    def download(self, file_id):
        self.sent.append(("download", {"file_id": file_id}))
        return hanjin_xlsx(*self.rows)


def _doc(name="한진 송장파일.xlsx"):
    return {"message": {"chat": {"id": OWNER}, "document": {"file_name": name, "file_id": "f1"}}}


def _invoice_bot(tmp_path, items, hanjin, dispatch=None, enabled=False):
    tg = InvoiceTelegram()
    tg.rows = hanjin
    ledger = tmp_path / "ledger.json"
    send = dispatch or (lambda batch: dispatch_or_simulate(batch, enabled=enabled))
    bot = OrderBot(tg, frozenset({OWNER}), pending_file=tmp_path / "p.json",
                   invoice_preparer=partial(prepare_invoices, fetch=lambda: items, ledger=ledger),
                   invoice_register=partial(register, dispatch=send, ledger=ledger),
                   invoice_pending_file=tmp_path / "inv.json")
    return tg, bot, ledger


def _buttons(tg):
    markup = tg.sent[-1][1]["reply_markup"]["inline_keyboard"]
    return [(row[0]["text"], row[0]["callback_data"]) for row in markup]


def _two_orders():
    items = [api_item("1"), api_item("2", "김철수")]
    hanjin = [hanjin_row(1, "111"), hanjin_row(2, "222", "김철수", **{"받으시는 분 전화": "01099999999"})]
    return items, hanjin


def test_document_gets_summary_files_and_buttons(tmp_path):
    tg, bot, _ = _invoice_bot(tmp_path, *_two_orders())
    bot.handle(_doc())
    assert [m for m, _ in tg.sent] == ["sendMessage", "download", "sendDocument", "sendDocument", "sendMessage"]
    assert tg.texts()[0] == "⏳ 송장 파일 확인 중…"
    labels = [text for text, _ in _buttons(tg)]
    assert labels == ["🚚 일치 1건 송장 등록", "🚚 확인 필요 1건 포함 (2건) 등록", "✖ 취소"]
    assert all(data.startswith("inv:") for _, data in _buttons(tg))


def test_nothing_to_register_means_no_buttons(tmp_path):
    tg, bot, _ = _invoice_bot(tmp_path, [api_item("1")], [hanjin_row(1, "111", "이영희")])
    bot.handle(_doc())
    assert "inline_keyboard" not in tg.sent[-1][1]["reply_markup"] and "주문 못 찾은 송장 1건" in tg.texts()[-1]


def test_non_excel_document_asks_for_a_tracking_file(tmp_path):
    tg, bot, _ = _invoice_bot(tmp_path, [], [])
    bot.handle(_doc("송장.pdf"))
    assert [m for m, _ in tg.sent] == ["sendMessage"] and ".xls" in tg.texts()[0]


def test_unreadable_document_is_reported(tmp_path):
    tg, bot, _ = _invoice_bot(tmp_path, [], [])
    tg.download = lambda file_id: b"junk"
    bot.handle(_doc())
    assert tg.texts()[-1].startswith("⚠️ 송장 확인 실패")


def test_matched_button_simulated_does_not_touch_the_ledger(tmp_path):
    tg, bot, ledger = _invoice_bot(tmp_path, *_two_orders())
    bot.handle(_doc())
    bot.handle(_press(_buttons(tg)[0][1]))
    assert "🧪 시험 모드" in tg.texts()[-1] and "등록했을 송장 1건" in tg.texts()[-1]
    assert not ledger.exists() and read_invoice_pending(tmp_path / "inv.json") == {}


def test_all_button_really_registers_and_records_only_successes(tmp_path):
    sent = []

    def dispatch(batch):
        sent.extend((i.product_order_id, i.tracking) for i in batch)
        return DispatchResult(("1",), (DispatchFailure("2", "E1", "이미 발송된 주문"),))

    tg, bot, ledger = _invoice_bot(tmp_path, *_two_orders(), dispatch=dispatch)
    bot.handle(_doc())
    bot.handle(_press(_buttons(tg)[1][1]))
    assert sent == [("1", "111"), ("2", "222")]
    text = tg.texts()[-1]
    assert "등록 완료 송장 1건" in text and "2 · 이미 발송된 주문" in text and "시험 모드" not in text
    assert list(read_invoice_ledger(ledger)) == ["1"]
    assert ("editMessageReplyMarkup", {"chat_id": OWNER, "message_id": 5, "reply_markup": {"inline_keyboard": []}}) in tg.sent


def test_cancel_and_expired_buttons(tmp_path):
    sent = []
    tg, bot, ledger = _invoice_bot(tmp_path, *_two_orders(), dispatch=lambda b: sent.append(b) or DispatchResult(()))
    bot.handle(_doc())
    cancel = _buttons(tg)[2][1]
    bot.handle(_press(cancel))
    assert tg.texts()[-1] == "취소했습니다. 등록하지 않았습니다."
    bot.handle(_press(cancel))
    assert "만료" in tg.texts()[-1] and sent == []


def test_invoice_button_survives_a_restart(tmp_path):
    tg, bot, _ = _invoice_bot(tmp_path, *_two_orders())
    bot.handle(_doc())
    data = _buttons(tg)[0][1]
    _, fresh, _ = _invoice_bot(tmp_path, [], [])
    fresh.tg = tg
    fresh.handle(_press(data))
    assert "송장 1건" in tg.texts()[-1]


def test_download_error_never_shows_the_token():
    from worker.telegram_bot import Telegram

    class Boom:
        def post(self, url, **kw):
            class R:
                def json(self):
                    return {"ok": True, "result": {"file_path": "documents/a.xlsx"}}
            return R()

        def get(self, url, **kw):
            import requests
            raise requests.ConnectionError(f"failed {url}")

    with pytest.raises(RuntimeError) as err:
        Telegram("SECRET-TOKEN", Boom()).download("f1")
    assert "SECRET-TOKEN" not in str(err.value) and err.value.__cause__ is None


def test_failure_message_names_what_failed():
    from worker.telegram_bot import failure_message

    assert failure_message(RuntimeError("x")).startswith("⚠️ 발주 처리 실패:")
    assert failure_message(RuntimeError("x"), "송장 등록").startswith("⚠️ 송장 등록 실패:")


# --- poll_forever relay timing -------------------------------------------------------------------

def test_poll_runs_relay_every_interval_and_survives_its_errors(caplog):
    from worker.telegram_bot import RELAY_INTERVAL, poll_forever

    clock = {"t": 1000.0}
    calls: list[float] = []

    def relay():
        calls.append(clock["t"])
        if len(calls) == 2:
            raise RuntimeError("github down")

    class _Tg:
        def __init__(self):
            self.n = 0

        def call(self, method, **params):
            self.n += 1
            clock["t"] += 10  # each long poll takes 10 fake seconds
            if self.n > 9:
                raise _Stop()
            return []

    bot = OrderBot(_Tg(), frozenset())
    with caplog.at_level("ERROR", logger="telegram_bot"), pytest.raises(_Stop):
        poll_forever(bot, sleep=lambda s: None, relay=relay, clock=lambda: clock["t"])
    assert calls[0] == 1000.0
    assert all(b - a >= RELAY_INTERVAL for a, b in zip(calls, calls[1:]))
    assert len(calls) == 4 and any("앱 송장 요청" in r.message for r in caplog.records)
