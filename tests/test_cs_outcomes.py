"""Outcome records: masking, data-repository write with outbox, and the 📊 CS 성적 numbers."""
import json
from datetime import datetime, timedelta

from store.base import ConflictError, StoreError
from store.memory_store import MemoryStore
from worker import cs_outcomes as co
from worker.cs_state import OUTCOMES_FILE, read_jsonl
from worker.naver_adapter import KST

NOW = datetime(2026, 10, 20, 12, 0, tzinfo=KST)
ENTRY = {"channel": "상품문의", "id": "5", "created": "2026-10-20T09:00:00+09:00", "product_name": "가짜 장어",
         "question": "010-1234-5678 로 연락 주세요 주문 1234567890123 언제 와요?",
         "draft": {"kind": "배송 일정", "route": "AI답변", "reason": "r", "draft": "내일 도착합니다 02-123-4567"}}


class BrokenStore(MemoryStore):
    def write_text(self, *args, **kwargs):
        raise StoreError("down")


class ConflictingStore(MemoryStore):
    """Another writer changes the file between our read and write, a few times."""

    def __init__(self, conflicts):
        super().__init__()
        self.conflicts = conflicts

    def write_text(self, path, content, expected_sha, message, author=None):
        if self.conflicts:
            self.conflicts -= 1
            raise ConflictError("changed")
        return super().write_text(path, content, expected_sha, message, author)


def outcome(key="qna:5", kind=co.AI_AS_IS, now=NOW, **kw):
    return co.build_outcome(key, ENTRY, kind, now, **kw)


def stored(store):
    snap = store.read_text(co.OUTCOMES_PATH)
    return json.loads(snap.content) if snap else None


# --- build + masking ----------------------------------------------------------------------------

def test_build_outcome_carries_route_kind_and_texts():
    o = outcome(kind=co.OWNER_WRITTEN, simulated=True, owner_answer="직접 쓴 답")
    assert (o["outcome"], o["route"], o["kind"], o["simulated"]) == ("owner_written", "AI답변", "배송 일정", True)
    assert o["ai_draft"].startswith("내일") and o["owner_answer"] == "직접 쓴 답" and o["settled_at"]
    assert o["channel"] == "상품문의" and set(o) >= {"id", "created", "question", "product_name"}


def test_mask_hides_phones_and_long_numbers_but_not_product_links():
    o = outcome(owner_answer="https://smartstore.naver.com/x/products/11368563340 보세요 010 9999 8888")
    masked = co.mask_outcome(o)
    assert "[전화]" in masked["question"] and "[번호]" in masked["question"]
    assert "5678" not in masked["question"] and "1234567890123" not in masked["question"]
    assert "02-123-4567" not in masked["ai_draft"]
    assert "products/11368563340" in masked["owner_answer"] and "[전화]" in masked["owner_answer"]
    assert o["question"].startswith("010-1234")  # the original is not changed


# --- record_outcome -----------------------------------------------------------------------------

def test_record_writes_local_unmasked_and_store_masked(tmp_path):
    store = MemoryStore()
    assert co.record_outcome(tmp_path, store, outcome()) is True
    (local,) = read_jsonl(tmp_path, OUTCOMES_FILE)
    assert "010-1234-5678" in local["question"]
    (remote,) = stored(store)
    assert "[전화]" in remote["question"] and "010-1234-5678" not in json.dumps(remote, ensure_ascii=False)


def test_same_outcome_twice_is_stored_once_and_list_is_trimmed(tmp_path, monkeypatch):
    store = MemoryStore()
    co.record_outcome(tmp_path, store, outcome())
    co.record_outcome(tmp_path, store, outcome())
    assert len(stored(store)) == 1
    monkeypatch.setattr(co, "MAX_STORED", 3)
    for n in range(5):
        co.record_outcome(tmp_path, store, outcome(key=f"qna:{n + 10}"))
    remote = stored(store)
    assert len(remote) == 3 and remote[-1]["id"] == co.short_id("qna:14")  # newest last


def test_write_conflicts_are_retried(tmp_path):
    store = ConflictingStore(conflicts=2)
    assert co.record_outcome(tmp_path, store, outcome()) is True and len(stored(store)) == 1


def test_store_failure_goes_to_outbox_and_does_not_raise(tmp_path, caplog):
    assert co.record_outcome(tmp_path, BrokenStore(), outcome()) is False
    waiting = json.loads((tmp_path / co.OUTBOX_FILE).read_text(encoding="utf-8"))
    assert len(waiting) == 1 and "010-1234" not in json.dumps(waiting, ensure_ascii=False)
    assert len(read_jsonl(tmp_path, OUTCOMES_FILE)) == 1
    assert "StoreError" in caplog.text


def test_no_store_still_records_locally(tmp_path):
    assert co.record_outcome(tmp_path, None, outcome()) is False
    assert len(read_jsonl(tmp_path, OUTCOMES_FILE)) == 1 and not (tmp_path / co.OUTBOX_FILE).exists()


def test_flush_sends_waiting_outcomes_once(tmp_path):
    co.record_outcome(tmp_path, BrokenStore(), outcome())
    store = MemoryStore()
    assert co.flush_outbox(tmp_path, store) == 1 and len(stored(store)) == 1
    assert not (tmp_path / co.OUTBOX_FILE).exists()
    assert co.flush_outbox(tmp_path, store) == 0


def test_flush_failure_keeps_outbox_and_never_raises(tmp_path):
    co.record_outcome(tmp_path, BrokenStore(), outcome())
    assert co.flush_outbox(tmp_path, BrokenStore()) == 0 and (tmp_path / co.OUTBOX_FILE).exists()


def test_broken_local_dir_does_not_raise(tmp_path):
    blocker = tmp_path / "file"
    blocker.write_text("x")
    assert co.record_outcome(blocker / "sub", None, outcome()) is False


# --- stats --------------------------------------------------------------------------------------

def test_stats_math_recent_and_all_time(tmp_path):
    from worker.cs_state import append_jsonl
    rows = [(co.AI_AS_IS, "AI답변", 1), (co.AI_AS_IS, "AI답변", 2), (co.AI_AS_IS, "사람에게", 3),
            (co.OWNER_WRITTEN, "AI답변", 1), (co.SELLER_CENTER, "사람에게", 0),
            (co.ANSWERED_ELSEWHERE, "AI답변", 1), (co.AI_AS_IS, "AI답변", 20), (co.OWNER_EDITED, "AI답변", 2)]
    for i, (kind, route, days) in enumerate(rows):
        entry = {**ENTRY, "draft": {**ENTRY["draft"], "route": route}}
        append_jsonl(tmp_path, OUTCOMES_FILE, co.build_outcome(f"qna:{i}", entry, kind, NOW - timedelta(days=days)))
    text = co.stats_text(tmp_path, NOW)
    recent, overall = text.split("\n\n")[1], text.split("\n\n")[2]
    assert "총 7건" in recent and "✅ AI 그대로 3건 (50%)" in recent      # 3 of 6 choices
    assert "✏️ 고쳐서 올림 1건 (17%)" in recent
    assert "✍️ 직접 씀 1건 (17%)" in recent and "🙋 판매자센터 1건 (17%)" in recent
    assert "따로 답함 1건" in recent
    assert "AI답변 분류 4건 중 ✅ 2건 (50%)" in recent and "사람에게 분류 2건 중 ✅ 1건 (50%)" in recent
    assert "총 8건" in overall and "✅ AI 그대로 4건 (57%)" in overall  # 4 of 7 choices
    assert text.rstrip().endswith("AI 그대로 비율이 충분히 높아지면 자동 등록을 검토할 수 있습니다.")


def test_stats_leave_out_questions_without_an_ai_draft(tmp_path):
    from worker.cs_state import append_jsonl
    no_draft = {**ENTRY, "draft": None}
    append_jsonl(tmp_path, OUTCOMES_FILE, co.build_outcome("qna:1", ENTRY, co.AI_AS_IS, NOW))
    append_jsonl(tmp_path, OUTCOMES_FILE, co.build_outcome("qna:2", no_draft, co.OWNER_WRITTEN, NOW))
    recent = co.stats_text(tmp_path, NOW).split("\n\n")[1]
    assert "총 2건" in recent and "✅ AI 그대로 1건 (100%)" in recent and "✍️ 직접 씀 0건 (0%)" in recent
    assert "AI 초안이 없던 1건" in recent


def test_stats_without_data_and_simulated_note(tmp_path):
    from worker.cs_state import append_jsonl
    assert "아직 없습니다" in co.stats_text(tmp_path, NOW)
    append_jsonl(tmp_path, OUTCOMES_FILE, outcome(simulated=True))
    assert "시험 모드" in co.stats_text(tmp_path, NOW)
