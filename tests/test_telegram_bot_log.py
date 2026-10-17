"""The bot's 송장 buttons write to the shared work log (and never let that break the reply)."""
from store import dispatch_queue as q
from store.memory_store import MemoryStore
from tests.test_telegram_bot import (
    DispatchFailure,
    DispatchResult,
    _buttons,
    _doc,
    _invoice_bot,
    _press,
    _two_orders,
)


def _bot_with_log(tmp_path, store, dispatch=None):
    tg, bot, _ledger = _invoice_bot(tmp_path, *_two_orders(), dispatch=dispatch)
    bot.log_store, bot.log_outbox = store, tmp_path / "outbox.json"
    bot.handle(_doc())
    return tg, bot


def _tap(bot, tg, index, first_name="민수"):
    press = _press(_buttons(tg)[index][1])
    press["callback_query"]["from"] = {"first_name": first_name}
    bot.handle(press)


def test_matched_button_logs_a_telegram_entry_with_simulated_flag(tmp_path):
    store = MemoryStore()
    tg, bot = _bot_with_log(tmp_path, store)
    _tap(bot, tg, 0)
    (entry,) = q.history(store)
    assert entry["source"] == "telegram" and entry["status"] == "done" and entry["by"] == "텔레그램 민수"
    assert entry["items"] == [{"ids": ["1"], "tracking": "111", "courier": "HANJIN", "recipient": "홍길동"}]
    assert entry["result"]["simulated"] is True and entry["result"]["success_ids"] == ["1"]
    assert q.pending(store) == [] and len(entry["id"]) == 12


def test_all_button_logs_every_shipment_and_real_results(tmp_path):
    store = MemoryStore()
    dispatch = lambda batch: DispatchResult(("1",), (DispatchFailure("2", "E1", "이미 발송된 주문"),))
    tg, bot = _bot_with_log(tmp_path, store, dispatch=dispatch)
    _tap(bot, tg, 1, first_name="")
    (entry,) = q.history(store)
    assert entry["by"] == "텔레그램" and [i["recipient"] for i in entry["items"]] == ["홍길동", "김철수"]
    assert entry["result"]["simulated"] is False
    assert entry["result"]["failures"] == [{"id": "2", "code": "E1", "message": "이미 발송된 주문"}]


def test_register_exception_logs_a_failed_entry_and_still_replies(tmp_path):
    store = MemoryStore()

    def boom(batch):
        raise RuntimeError("GW.IP_NOT_ALLOWED " + "x" * 400)

    tg, bot = _bot_with_log(tmp_path, store, dispatch=boom)
    _tap(bot, tg, 0)
    assert tg.texts()[-1].startswith("⚠️ 송장 등록 실패")
    (entry,) = q.history(store)
    assert entry["status"] == "failed" and entry["result"]["error"].startswith("RuntimeError: GW.IP_NOT_ALLOWED")
    assert len(entry["result"]["error"]) <= len("RuntimeError: ") + 300


class _BrokenStore(MemoryStore):
    def read_text(self, path):
        raise RuntimeError("store exploded")


def test_bot_still_replies_and_clears_pending_when_the_log_store_raises(tmp_path):
    tg, bot = _bot_with_log(tmp_path, _BrokenStore())
    _tap(bot, tg, 0)
    assert "🧪 시험 모드" in tg.texts()[-1]
    assert bot.invoice_pending_file.read_text(encoding="utf-8").strip() == "{}"
    assert (tmp_path / "outbox.json").exists()  # kept for the relay loop to retry


def test_cancel_is_not_logged_and_no_store_is_fine(tmp_path):
    store = MemoryStore()
    tg, bot = _bot_with_log(tmp_path, store)
    _tap(bot, tg, 2)
    assert tg.texts()[-1] == "취소했습니다. 등록하지 않았습니다." and q.history(store) == []
    tg2, bot2 = _bot_with_log(tmp_path / "other", None)
    _tap(bot2, tg2, 0)
    assert "🧪 시험 모드" in tg2.texts()[-1]
