"""store.dispatch_queue with synthetic data on MemoryStore."""
import json
from datetime import datetime

import pytest

from core.invoice_match import KST
from store import dispatch_queue as q
from store.base import ConflictError
from store.memory_store import MemoryStore

NOW = datetime(2026, 10, 13, 9, 0, tzinfo=KST)


def _item(pid="1001", tracking="111", courier="HANJIN"):
    return {"ids": [pid], "tracking": tracking, "courier": courier, "recipient": "홍길동"}


def test_submit_pending_get_complete():
    store = MemoryStore()
    rid = q.submit(store, [_item()], "민수", NOW)
    assert len(rid) == 12 and int(rid, 16) >= 0
    request = q.get(store, rid)
    assert request["status"] == "pending" and request["by"] == "민수" and request["result"] is None
    assert request["created_at"] == "2026-10-13T09:00:00+09:00"
    assert [r["id"] for r in q.pending(store)] == [rid]
    result = {"processed_at": "x", "simulated": False, "success_ids": ["1001"], "failures": [], "error": None}
    q.complete(store, rid, result, "done")
    assert q.pending(store) == [] and q.get(store, rid)["result"] == result
    assert q.get(store, "nope") is None


def test_complete_unknown_request_and_bad_status():
    store = MemoryStore()
    with pytest.raises(KeyError):
        q.complete(store, "missing", {}, "done")
    rid = q.submit(store, [_item()], "a", NOW)
    with pytest.raises(ValueError):
        q.complete(store, rid, {}, "pending")


def test_empty_queue_reads_empty():
    assert q.pending(MemoryStore()) == []


class _ConflictOnce(MemoryStore):
    def __init__(self):
        super().__init__()
        self.fail_next = False

    def write_text(self, path, content, expected_sha, message, author=None):
        if self.fail_next:
            self.fail_next = False
            raise ConflictError("raced")
        return super().write_text(path, content, expected_sha, message, author)


def test_conflict_is_retried_and_gives_up_after_three():
    store = _ConflictOnce()
    store.fail_next = True
    rid = q.submit(store, [_item()], "a", NOW)
    assert q.get(store, rid) is not None
    store.fail_next = True
    q.complete(store, rid, {"success_ids": []}, "done")
    assert q.get(store, rid)["status"] == "done"

    class _Always(MemoryStore):
        def write_text(self, *a, **k):
            raise ConflictError("always")

    with pytest.raises(ConflictError):
        q.submit(_Always(), [_item()], "a", NOW)


def test_duplicate_submit_returns_existing_id():
    store = MemoryStore()
    first = q.submit(store, [_item("1"), _item("2", "222")], "a", NOW)
    assert q.submit(store, [_item("2", "222"), _item("1")], "b", NOW) == first
    assert len(json.loads(store.read_text(q.DISPATCH_QUEUE_PATH).content)["requests"]) == 1
    q.complete(store, first, {"simulated": False, "failures": []}, "done")
    assert q.submit(store, [_item("1"), _item("2", "222")], "a", NOW) == first  # really done: no re-send


def test_failed_or_simulated_requests_can_be_resubmitted():
    store = MemoryStore()
    rid = q.submit(store, [_item()], "a", NOW)
    q.complete(store, rid, {"error": "x"}, "failed")
    second = q.submit(store, [_item()], "a", NOW)
    assert second != rid
    q.complete(store, second, {"simulated": True, "failures": []}, "done")
    assert q.submit(store, [_item()], "a", NOW) not in (rid, second)


def test_trim_keeps_max_and_drops_oldest_finished():
    store = MemoryStore()
    ids = []
    for n in range(q.MAX_REQUESTS + 4):
        rid = q.submit(store, [_item(str(1000 + n), str(n))], "a", NOW)
        ids.append(rid)
        if n % 2 == 0:
            q.complete(store, rid, {"failures": []}, "done")
    requests = json.loads(store.read_text(q.DISPATCH_QUEUE_PATH).content)["requests"]
    assert len(requests) == q.MAX_REQUESTS
    kept = [r["id"] for r in requests]
    assert ids[0] not in kept and ids[-1] in kept and ids[1] in kept  # oldest done dropped, pending kept


@pytest.mark.parametrize("items, message", [
    ([], "없습니다"),
    ([_item("12a")], "숫자"),
    ([{**_item(), "ids": []}], "숫자"),
    ([_item(tracking="1 2")], "송장번호"),
    ([_item(tracking="x" * 51)], "송장번호"),
    ([_item(courier="FEDEX")], "택배사"),
    ([_item(str(n)) for n in range(501)], "500"),
])
def test_validation_errors(items, message):
    with pytest.raises(ValueError, match=message):
        q.validate_items(items)
    with pytest.raises(ValueError):
        q.submit(MemoryStore(), items, "a", NOW)


def test_validation_accepts_all_known_couriers():
    from core.invoice_match import COURIER_CODES
    q.validate_items([_item(courier=c, tracking="AB-12") for c in set(COURIER_CODES.values())])
