"""cs_examples: the owner's own past answers become examples in the next AI draft prompt."""
import json

from worker import cs_examples
from worker.cs_examples import Example, examples_text, load_examples, pick_examples
from worker.cs_state import OUTCOMES_FILE, PAIRS_FILE


def _write(path, records):
    path.write_text("".join(json.dumps(r, ensure_ascii=False) + "\n" for r in records), encoding="utf-8")


def test_owner_answers_come_from_outcomes_and_pairs_newest_first(tmp_path):
    _write(tmp_path / OUTCOMES_FILE, [
        {"outcome": "owner_edited", "settled_at": "2026-10-26T10:00:00+09:00", "product_name": "장어",
         "question": "언제 와요?", "owner_answer": "내일 도착합니다^^"},
        {"outcome": "ai_as_is", "settled_at": "2026-10-26T11:00:00+09:00", "product_name": "장어",
         "question": "AI 그대로", "owner_answer": ""},  # no owner text: not an example
        {"outcome": "answered_elsewhere", "settled_at": "2026-10-27T09:00:00+09:00", "product_name": "가리비",
         "question": "보관은요?", "owner_answer": "냉장 보관해 주세요"},
    ])
    _write(tmp_path / PAIRS_FILE, [
        {"answered_seen_at": "2026-10-26T12:00:00+09:00", "question": "포장 돼요?", "owner_answer": "네 가능합니다"},
        {"answered_seen_at": "2026-10-27T09:10:00+09:00", "question": "보관은요?", "owner_answer": "냉장 보관해 주세요"},
    ])
    examples = load_examples(tmp_path)
    assert [e.question for e in examples] == ["보관은요?", "포장 돼요?", "언제 와요?"]  # duplicate dropped
    assert examples[0] == Example("보관은요?", "냉장 보관해 주세요", "가리비", "2026-10-27T09:00:00+09:00")


def test_no_files_no_examples(tmp_path):
    assert load_examples(tmp_path) == []


def test_answers_promising_a_follow_up_are_not_examples(tmp_path):
    _write(tmp_path / OUTCOMES_FILE, [
        {"outcome": "owner_written", "settled_at": "2026-10-26T10:00:00+09:00", "product_name": "장어",
         "question": "재고 있어요?", "owner_answer": "확인 후 연락드릴게요"}])
    assert load_examples(tmp_path) == []


def test_broken_lines_are_skipped_not_fatal(tmp_path):
    (tmp_path / OUTCOMES_FILE).write_text(
        'not json\n{"outcome": "owner_written", "settled_at": "t", "question": "q", "owner_answer": "a"}\n',
        encoding="utf-8")
    assert [e.answer for e in load_examples(tmp_path)] == ["a"]


def test_pick_prefers_the_same_product_then_recency():
    history = [Example(f"q{i}", f"a{i}", "가리비" if i % 2 else "장어", f"t{9 - i}") for i in range(8)]
    picked = pick_examples(history, "장어", limit=3)
    assert [e.question for e in picked] == ["q0", "q2", "q4"]
    assert [e.question for e in pick_examples(history, "", limit=2)] == ["q0", "q1"]


def test_text_is_masked_trimmed_and_says_how_to_use_it():
    long = "가" * (cs_examples.ANSWER_MAX + 50)
    text = examples_text([Example("주문 2026100700001 언제 와요? 010-1234-5678", long, "장어", "t")])
    assert "[번호]" in text and "[전화]" in text and "010-1234" not in text
    assert "가" * cs_examples.ANSWER_MAX + "…" in text
    assert "말투" in text and "그날 상황" in text
    assert examples_text([]) == ""
