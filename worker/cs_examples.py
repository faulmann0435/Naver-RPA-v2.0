"""The owner's own past answers as examples for the next AI draft, so drafts follow the owner over time.

Sources (CS state dir, this PC): outcomes.jsonl (answers posted from Telegram, or found answered in
the seller center) and pairs.jsonl (alerted questions the owner answered in the seller center).
Only the owner's text counts; AI drafts posted unchanged add nothing new. Answers that promise to get
back to the customer are skipped, like in the drafts themselves (worker.cs_drafter.FOLLOW_UP).
Examples guide tone and shape only: the store document stays the source of facts.
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from pathlib import Path

from worker.cs_drafter import FOLLOW_UP, mask_question
from worker.cs_state import OUTCOMES_FILE, PAIRS_FILE

log = logging.getLogger("cs_examples")

MAX_EXAMPLES = 8  # per draft; enough to show the tone without crowding out the question
QUESTION_MAX, ANSWER_MAX = 200, 400
GUIDE = ("[사장님 최근 답변 예시]\n"
         "사장님이 최근 직접 쓰거나 고쳐서 올린 답변입니다. 말투·길이·구성을 참고하세요.\n"
         "날짜·재고·조업·입고처럼 그날 상황에 따른 내용은 따라 하지 마세요.\n"
         "가게 정보 문서와 다르면 문서를 따르고, 다른 점을 reason에 적으세요.")


@dataclass(frozen=True)
class Example:
    question: str
    answer: str
    product_name: str
    when: str  # ISO time the answer was seen; newest first


def _records(path: Path) -> list[dict]:
    if not path.exists():
        return []
    records = []
    for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip():
            continue
        try:
            records.append(json.loads(line))
        except ValueError:
            log.warning("%s %d번째 줄을 읽지 못해 건너뜁니다.", path.name, number)
    return records


def _example(record: dict, when_key: str) -> Example | None:
    question = str(record.get("question") or "").strip()
    answer = str(record.get("owner_answer") or "").strip()
    if not question or not answer or FOLLOW_UP.search(answer):
        return None
    return Example(question, answer, str(record.get("product_name") or ""), str(record.get(when_key) or ""))


def load_examples(state_dir: Path) -> list[Example]:
    """Every usable owner answer, newest first, one per question (outcomes win over pairs)."""
    found: dict[str, Example] = {}
    sources = ((OUTCOMES_FILE, "settled_at"), (PAIRS_FILE, "answered_seen_at"))
    for name, when_key in sources:
        for record in _records(state_dir / name):
            example = _example(record, when_key)
            if example and example.question not in found:
                found[example.question] = example
    return sorted(found.values(), key=lambda e: e.when, reverse=True)


def pick_examples(history: list[Example], product_name: str, limit: int = MAX_EXAMPLES) -> list[Example]:
    """Answers about the same product first, then the newest others."""
    same = [e for e in history if product_name and e.product_name == product_name]
    others = [e for e in history if e not in same]
    return (same + others)[:limit]


def _cut(text: str, limit: int) -> str:
    text = mask_question(text)
    return text if len(text) <= limit else text[:limit] + "…"


def examples_text(examples: list[Example]) -> str:
    """The prompt block; '' when there is nothing to show."""
    if not examples:
        return ""
    blocks = [f"예시 {n}. 상품: {e.product_name or '(알 수 없음)'}\n"
              f"고객: {_cut(e.question, QUESTION_MAX)}\n사장님: {_cut(e.answer, ANSWER_MAX)}"
              for n, e in enumerate(examples, 1)]
    return "\n\n".join([GUIDE, *blocks])
