"""Telegram remote control: who may use it, what a run sends, and the 공유 완료 button."""
import io

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


def test_stranger_only_learns_their_chat_id():
    tg, calls = FakeTelegram(), []
    bot = OrderBot(tg, frozenset({OWNER}), runner=lambda: calls.append(1) or _run())
    bot.handle(_message(999, RUN_BUTTON))
    assert calls == [] and "999" in tg.texts()[0]


def test_run_sends_files_then_summary_with_done_button():
    tg = FakeTelegram()
    bot = OrderBot(tg, frozenset({OWNER}), runner=_run)
    bot.handle(_message(OWNER, RUN_BUTTON))
    methods = [m for m, _ in tg.sent]
    assert methods == ["sendMessage", "sendDocument", "sendMessage"]
    last = tg.sent[-1][1]
    assert last["text"].startswith("📦 발주서 1개")
    assert last["reply_markup"]["inline_keyboard"][0][0]["callback_data"].startswith("done:")


def test_done_button_records_orders_once():
    tg, marked = FakeTelegram(), []
    bot = OrderBot(tg, frozenset({OWNER}), runner=_run, mark=lambda ids: marked.append(ids) or len(ids))
    bot.handle(_message(OWNER, RUN_BUTTON))
    data = tg.sent[-1][1]["reply_markup"]["inline_keyboard"][0][0]["callback_data"]
    press = {"callback_query": {"id": "q", "data": data, "message": {"chat": {"id": OWNER}, "message_id": 5}}}
    bot.handle(press)
    bot.handle(press)
    assert marked == [("1", "2")]
    assert "2건을 공유 완료" in tg.texts()[-2] and "만료" in tg.texts()[-1]


def test_failure_is_reported_not_swallowed():
    tg = FakeTelegram()

    def broken():
        raise RuntimeError("HTTP 403: GW.IP_NOT_ALLOWED")

    OrderBot(tg, frozenset({OWNER}), runner=broken).handle(_message(OWNER, RUN_BUTTON))
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
