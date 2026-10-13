"""worker.dispatch_relay with fake register/notify and MemoryStore."""
import json
from datetime import datetime

from core.invoice_match import KST
from store import dispatch_queue as q
from store.base import StoreError
from store.memory_store import MemoryStore
from worker import dispatch_relay as relay
from worker.dispatch_client import DispatchFailure, DispatchResult

NOW = datetime(2026, 10, 13, 9, 0, tzinfo=KST)


def _item(pid="1001", tracking="111", courier="HANJIN"):
    return {"ids": [pid], "tracking": tracking, "courier": courier, "recipient": "홍길동"}


class _Register:
    def __init__(self, result=None, error=None):
        self.result, self.error, self.calls = result, error, []

    def __call__(self, entry, include_checks):
        self.calls.append((entry, include_checks))
        if self.error:
            raise self.error
        return self.result


def _run(store, register, tmp_path):
    notes: list[str] = []
    count = relay.process_queue(store, register_fn=register, notify=notes.append,
                                now=lambda: NOW, processed_file=tmp_path / "p.json")
    return count, notes


def test_pending_request_is_registered_and_completed(tmp_path):
    store = MemoryStore()
    rid = q.submit(store, [_item(), _item("1002", "222", "CJGLS")], "민수", NOW)
    register = _Register(DispatchResult(("1001",), (DispatchFailure("1002", "E", "주소 오류"),)))
    count, notes = _run(store, register, tmp_path)
    assert count == 1
    entry, include = register.calls[0]
    assert include is False and entry["includable"] == [] and entry["matched"] == q.get(store, rid)["items"]
    done = q.get(store, rid)
    assert done["status"] == "done"
    assert done["result"]["success_ids"] == ["1001"] and done["result"]["error"] is None
    assert done["result"]["failures"] == [{"id": "1002", "code": "E", "message": "주소 오류"}]
    assert len(notes) == 1 and notes[0].startswith("🚀 앱에서 요청한 송장 등록 (민수)")
    assert "등록 완료 송장 1건" in notes[0] and "1002 · 주소 오류" in notes[0]


def test_simulated_result_passes_through(tmp_path):
    store = MemoryStore()
    rid = q.submit(store, [_item()], "a", NOW)
    _, notes = _run(store, _Register(DispatchResult(("1001",), simulated=True)), tmp_path)
    assert q.get(store, rid)["result"]["simulated"] is True
    assert "시험 모드" in notes[0]


def test_oldest_first_and_nothing_pending(tmp_path):
    store = MemoryStore()
    q.submit(store, [_item("1")], "a", NOW)
    q.submit(store, [_item("2", "222")], "b", NOW)
    register = _Register(DispatchResult(()))
    assert _run(store, register, tmp_path)[0] == 2
    assert [c[0]["matched"][0]["ids"] for c in register.calls] == [["1"], ["2"]]
    assert q.pending(store) == []
    assert _run(store, register, tmp_path)[0] == 0


def test_exception_marks_failed_and_is_never_retried(tmp_path):
    store = MemoryStore()
    rid = q.submit(store, [_item()], "a", NOW)
    register = _Register(error=RuntimeError("IP_NOT_ALLOWED"))
    _, notes = _run(store, register, tmp_path)
    failed = q.get(store, rid)
    assert failed["status"] == "failed" and "IP_NOT_ALLOWED" in failed["result"]["error"]
    assert "등록 실패" in notes[0]
    _run(store, register, tmp_path)
    assert len(register.calls) == 1


def test_invalid_items_fail_without_dispatch(tmp_path):
    store = MemoryStore()
    bad = [{**_item(), "tracking": "bad tracking!"}]
    seeded = {"requests": [{"id": "abc", "created_at": "t", "by": "x", "status": "pending",
                            "items": bad, "result": None}]}
    store.write_text(q.DISPATCH_QUEUE_PATH, json.dumps(seeded), None, "seed")
    register = _Register(DispatchResult(()))
    _run(store, register, tmp_path)
    assert register.calls == [] and q.get(store, "abc")["status"] == "failed"


class _NoWrite(MemoryStore):
    blocked = False

    def write_text(self, path, content, expected_sha, message, author=None):
        if self.blocked:
            raise StoreError("GitHub 저장 실패 (HTTP 403): forbidden")
        return super().write_text(path, content, expected_sha, message, author)


def test_locally_processed_request_is_skipped_even_if_still_pending(tmp_path):
    store = _NoWrite()
    rid = q.submit(store, [_item()], "a", NOW)
    store.blocked = True
    register = _Register(DispatchResult(("1001",)))
    _, notes = _run(store, register, tmp_path)
    assert len(register.calls) == 1 and q.get(store, rid)["status"] == "pending"
    assert notes[0] == relay.WRITE_FAILED_NOTICE and "등록 완료" in notes[1]
    _, again = _run(store, register, tmp_path)  # next round: still pending in the queue
    assert len(register.calls) == 1 and again == []
    store.blocked = False  # permission fixed: the stored local result is saved now
    _run(store, register, tmp_path)
    assert len(register.calls) == 1 and q.get(store, rid)["status"] == "done"


def test_interrupted_request_is_failed_not_resent(tmp_path):
    store = MemoryStore()
    rid = q.submit(store, [_item()], "a", NOW)
    relay._remember(tmp_path / "p.json", rid, {"status": relay.STARTED})
    register = _Register(DispatchResult(()))
    _run(store, register, tmp_path)
    assert register.calls == [] and q.get(store, rid)["status"] == "failed"


def test_notify_failure_does_not_break(tmp_path):
    store = MemoryStore()
    rid = q.submit(store, [_item()], "a", NOW)

    def boom(text):
        raise RuntimeError("telegram down")

    relay.process_queue(store, register_fn=_Register(DispatchResult(("1001",))), notify=boom,
                        now=lambda: NOW, processed_file=tmp_path / "p.json")
    assert q.get(store, rid)["status"] == "done"


def test_load_store(tmp_path):
    assert relay.load_store(tmp_path / "none.toml") is None
    f = tmp_path / "s.toml"
    f.write_text("[other]\na = 1\n", encoding="utf-8")
    assert relay.load_store(f) is None
    f.write_text('[data_store]\nrepo = "o/r"\nbranch = "dev"\ntoken = "fake-token"\n', encoding="utf-8")
    assert relay.load_store(f).repo == "o/r"
