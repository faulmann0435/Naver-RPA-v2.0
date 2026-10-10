"""Telegram remote control: who may use it, what a run sends, and the 공유 완료 button."""
import io

import pytest

from worker.order_run import RunResult, mark_shared, read_ledger, run_orders, summary
from worker.telegram_bot import RUN_BUTTON, OrderBot

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
        "quantity": 1, "packageNumber": "7", "productId": "5", "shippingAddress": {"zipCode": "12345"}}}
    engine = lambda: type("E", (), {"config": None, "dictionary": None, "settings": None, "source": "t"})()
    mark_shared(("1",), ledger)
    run = run_orders(fetch=lambda: [item("1")], engine_loader=engine, ledger=ledger)
    assert run.paid == 0 and run.already_shared == 1 and run.files == []
    assert "새로 발주할 주문이 없습니다" in summary(run)
    assert read_ledger(ledger) == {"1"}


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
