"""Shared 송장 work log: queue source/record/history, entry_items, and worker.dispatch_log."""
import json
import logging
from datetime import datetime

import pytest

from core.invoice_match import KST
from store import dispatch_queue as q
from store.base import ConflictError, StoreError
from store.memory_store import MemoryStore
from worker import dispatch_log as dl
from worker import dispatch_relay as relay
from worker.dispatch_client import DispatchFailure, DispatchResult
from worker.invoice_run import entry_items

NOW = datetime(2026, 10, 13, 9, 0, tzinfo=KST)
OLDER = datetime(2026, 10, 12, 9, 0, tzinfo=KST)


def _item(pid="1001", tracking="111", courier="HANJIN"):
    return {"ids": [pid], "tracking": tracking, "courier": courier, "recipient": "홍길동"}


ITEMS = [_item()]
RESULT = {"processed_at": "x", "simulated": False, "success_ids": ["1001"], "failures": [], "error": None}


# --- store.dispatch_queue: source, record, history -----------------------------------------------------

def test_submit_sets_source_app_and_legacy_entries_count_as_app():
    store = MemoryStore()
    rid = q.submit(store, ITEMS, "a", NOW)
    assert q.get(store, rid)["source"] == q.SOURCE_APP == "app"
    assert q.source_of({"id": "old"}) == "app" and q.source_of({"source": "telegram"}) == "telegram"


def test_record_appends_finished_telegram_entry():
    store = MemoryStore()
    q.record(store, "abc123", ITEMS, "텔레그램 민수", NOW, RESULT, "done")
    entry = q.get(store, "abc123")
    assert entry["source"] == "telegram" and entry["status"] == "done" and entry["by"] == "텔레그램 민수"
    assert entry["items"] == ITEMS and entry["result"] == RESULT
    assert entry["created_at"] == "2026-10-13T09:00:00+09:00"


def test_record_cleans_items_like_submit():
    store = MemoryStore()
    q.record(store, "x", [{"ids": ["1"], "tracking": 5, "courier": "CJGLS", "extra": 1}], "a", NOW, RESULT, "failed")
    assert q.get(store, "x")["items"] == [{"ids": ["1"], "tracking": "5", "courier": "CJGLS", "recipient": ""}]


def test_record_is_idempotent_and_rejects_pending():
    store = MemoryStore()
    q.record(store, "abc", ITEMS, "a", NOW, RESULT, "done")
    q.record(store, "abc", [_item("2")], "b", NOW, {"error": "x"}, "failed")  # a retry: no duplicate, no change
    requests = json.loads(store.read_text(q.DISPATCH_QUEUE_PATH).content)["requests"]
    assert len(requests) == 1 and requests[0]["status"] == "done" and requests[0]["items"][0]["ids"] == ["1001"]
    with pytest.raises(ValueError):
        q.record(store, "new", ITEMS, "a", NOW, RESULT, "pending")
    assert q.get(store, "new") is None


class _ConflictOnce(MemoryStore):
    fail_next = False

    def write_text(self, path, content, expected_sha, message, author=None):
        if self.fail_next:
            self.fail_next = False
            raise ConflictError("raced")
        return super().write_text(path, content, expected_sha, message, author)


def test_record_retries_on_conflict():
    store = _ConflictOnce()
    store.fail_next = True
    q.record(store, "abc", ITEMS, "a", NOW, RESULT, "failed")
    assert q.get(store, "abc")["status"] == "failed"


def test_history_is_newest_first_and_pending_ignores_recorded():
    store = MemoryStore()
    q.record(store, "old", ITEMS, "a", OLDER, RESULT, "done")
    mine = q.submit(store, [_item("2", "222")], "a", NOW)
    q.record(store, "same-time", [_item("3", "333")], "a", NOW, RESULT, "done")
    assert [r["id"] for r in q.history(store)] == ["same-time", mine, "old"]
    assert [r["id"] for r in q.pending(store)] == [mine]
    assert q.history(MemoryStore()) == []


def test_recorded_entry_is_never_picked_up_by_the_relay(tmp_path):
    store = MemoryStore()
    q.record(store, "tg", ITEMS, "a", NOW, RESULT, "done")
    calls = []
    count = relay.process_queue(store, register_fn=lambda *a: calls.append(a), processed_file=tmp_path / "p.json")
    assert count == 0 and calls == []


def test_trim_applies_to_recorded_entries_too():
    store = MemoryStore()
    for n in range(q.MAX_REQUESTS + 2):
        q.record(store, f"r{n}", [_item(str(1000 + n), str(n))], "a", NOW, RESULT, "done")
    kept = [r["id"] for r in q.history(store)]
    assert q.MAX_REQUESTS == 300 and len(kept) == 300 and "r0" not in kept and "r301" in kept


# --- worker.invoice_run.entry_items ----------------------------------------------------------------------

def test_entry_items_for_new_and_old_pending_formats():
    new = {"when": "x", "matched": [_item()],
           "includable": [_item(), {"ids": ["2", "3"], "tracking": "222", "courier": "CJGLS", "recipient": "김철수"}]}
    assert entry_items(new, False) == [_item()]
    assert [i["recipient"] for i in entry_items(new, True)] == ["홍길동", "김철수"]
    no_recipient = {"matched": [{"ids": ["1"], "tracking": "111", "courier": "EPOST"}], "includable": []}
    assert entry_items(no_recipient, False) == [{"ids": ["1"], "tracking": "111", "courier": "EPOST", "recipient": ""}]
    old = {"matched": [["1", "2", "111"]], "includable": [["1", "2", "111"], ["3", "222"]]}
    assert entry_items(old, True) == [
        {"ids": ["1", "2"], "tracking": "111", "courier": "HANJIN", "recipient": ""},
        {"ids": ["3"], "tracking": "222", "courier": "HANJIN", "recipient": ""}]
    q.validate_items(entry_items(old, True))


# --- worker.dispatch_log -----------------------------------------------------------------------------------

class _Down(MemoryStore):
    down = True

    def write_text(self, path, content, expected_sha, message, author=None):
        if self.down:
            raise StoreError("GitHub 저장 실패: https://x/?token=SECRET-TOKEN")
        return super().write_text(path, content, expected_sha, message, author)


def _log(store, tmp_path, rid="r1"):
    return dl.log_run(store, rid, ITEMS, "텔레그램 민수", NOW, RESULT, "done", tmp_path / "outbox.json")


def test_result_shapes():
    result = DispatchResult(("1",), (DispatchFailure("2", "E", "주소 오류"),), simulated=True)
    assert dl.result_dict(result, NOW) == {"processed_at": "2026-10-13T09:00:00+09:00", "simulated": True,
                                           "success_ids": ["1"], "error": None,
                                           "failures": [{"id": "2", "code": "E", "message": "주소 오류"}]}
    assert dl.failed_dict("boom", NOW)["error"] == "boom"


def test_log_run_reaches_the_store(tmp_path):
    store = MemoryStore()
    assert _log(store, tmp_path) is True
    entry = q.get(store, "r1")
    assert entry["source"] == "telegram" and entry["by"] == "텔레그램 민수" and entry["items"] == ITEMS
    assert not (tmp_path / "outbox.json").exists()


def test_log_run_without_store_does_nothing(tmp_path):
    assert _log(None, tmp_path) is False
    assert not (tmp_path / "outbox.json").exists()


def test_store_failure_goes_to_outbox_without_leaking_the_message(tmp_path, caplog):
    with caplog.at_level("ERROR", logger="telegram_bot"):
        assert _log(_Down(), tmp_path) is False
    waiting = json.loads((tmp_path / "outbox.json").read_text(encoding="utf-8"))
    assert list(waiting) == ["r1"] and waiting["r1"]["status"] == "done"
    assert "StoreError" in caplog.text and "SECRET-TOKEN" not in caplog.text


def test_flush_sends_waiting_entries_and_keeps_failing_ones(tmp_path):
    outbox, store = tmp_path / "outbox.json", _Down()
    _log(store, tmp_path, "a")
    _log(store, tmp_path, "b")
    assert dl.flush_outbox(store, outbox) == 0
    assert set(json.loads(outbox.read_text(encoding="utf-8"))) == {"a", "b"}
    store.down = False
    assert dl.flush_outbox(store, outbox) == 2
    assert json.loads(outbox.read_text(encoding="utf-8")) == {}
    assert {r["id"] for r in q.history(store)} == {"a", "b"}
    assert dl.flush_outbox(store, outbox) == 0


def test_flush_survives_one_bad_entry_and_never_duplicates(tmp_path):
    outbox, store = tmp_path / "outbox.json", MemoryStore()
    q.record(store, "done-before", ITEMS, "x", NOW, RESULT, "done")  # a partly successful earlier attempt
    good = {"items": ITEMS, "by": "x", "created_at": NOW.isoformat(), "result": RESULT, "status": "done"}
    outbox.write_text(json.dumps({"bad": {"items": ITEMS}, "done-before": good, "good": good}), encoding="utf-8")
    assert dl.flush_outbox(store, outbox) == 2
    assert list(json.loads(outbox.read_text(encoding="utf-8"))) == ["bad"]
    assert sorted(r["id"] for r in q.history(store)) == ["done-before", "good"]


def test_unreadable_outbox_is_not_fatal(tmp_path):
    outbox = tmp_path / "outbox.json"
    outbox.write_text("not json", encoding="utf-8")
    assert dl.flush_outbox(MemoryStore(), outbox) == 0


def test_relay_flushes_the_outbox_and_survives_its_errors(tmp_path, caplog, monkeypatch):
    store, outbox = _Down(), tmp_path / "outbox.json"
    _log(store, tmp_path)
    store.down = False
    relay.make_relay(store, lambda text: None, outbox)()
    assert q.get(store, "r1") is not None

    def boom(*a, **k):
        raise RuntimeError("x")

    monkeypatch.setattr(relay, "flush_outbox", boom)
    with caplog.at_level(logging.ERROR, logger="telegram_bot"):
        relay.make_relay(store, lambda text: None, outbox)()
    assert "미전송 작업 기록 처리 실패" in caplog.text
