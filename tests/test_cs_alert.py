"""Alert text: layout per link kind, no-AI mode, claims, trimming, no personal values."""
from datetime import datetime

import pytest

from tests.test_order_link import PERSONAL, detail, qna
from worker.cs_alert import NO_AI_NOTE, TELEGRAM_LIMIT, format_alert, trim
from worker.cs_drafter import DraftResult, InquiryView
from worker.naver_adapter import KST
from worker.order_link import LinkResult, link_customer_inquiry, link_qna

ASKED = datetime(2026, 10, 10, 9, 5, tzinfo=KST)
AI = DraftResult("배송 일정", "AI답변", "문서에 있음", "안녕하세요 고객님^^ 내일 도착합니다.")
HUMAN = DraftResult("취소·변경", "사람에게", "주소 변경 요청", "확인 후 안내드리겠습니다")


def make(channel="상품문의", question="언제 와요?", name="국산 풍천 장어", link=None):
    link = link or link_customer_inquiry({"productOrderIdList": "2026100700001-1"}, [detail()])
    return InquiryView(channel, ASKED, name, question, link), link


def test_exact_link_with_ai_draft_layout():
    view, link = make()
    lines = format_alert(view, link, AI).split("\n")
    assert lines[0] == "🟡 상품문의 · 국산 풍천 장어"
    assert lines[1] == "2026-10-10(토) 09:05"
    assert lines[2] == '고객: "언제 와요?"'
    assert lines[3] == "주문: 정확"
    assert lines[4] == "- 2026-10-07 12:42 · 구매확정 · 1kg ×2"
    assert lines[5] == "분류: 배송 일정 · AI 초안"
    assert lines[6] == "✍️ AI 초안:" and lines[7] == AI.draft
    assert "판매자센터" not in "\n".join(lines)


def test_inquiry_uses_blue_icon():
    view, link = make(channel="고객문의")
    assert format_alert(view, link, AI).startswith("🔵 고객문의")


def test_claim_uses_red_icon():
    view, link = make()
    claim = DraftResult("불만·클레임", "사람에게", "상함 신고", "사과드립니다")
    assert format_alert(view, link, claim).startswith("🔴 상품문의")


def test_human_route_shows_reason():
    view, link = make()
    assert "분류: 취소·변경 · 사장님 직접 처리 권장 — 주소 변경 요청" in format_alert(view, link, HUMAN)


def test_guess_and_many_add_hint_and_none_lists_no_orders():
    guess = link_qna(qna(), [detail()])
    many = link_qna(qna(), [detail(), detail(order_id="2", po_id="2-1")])
    none = LinkResult("없음", (), "x")
    for link, kind in ((guess, "추정"), (many, "후보 여러 건")):
        text = format_alert(make(link=link)[0], link, AI)
        assert f"주문: {kind}" in text and "판매자센터에서 작성자 아이디를 눌러" in text
    text = format_alert(make(link=none)[0], none, AI)
    assert "주문: 없음\n분류" in text and "판매자센터" not in text


def test_no_ai_mode_and_failure_note():
    view, link = make()
    text = format_alert(view, link, None)
    assert text.endswith(NO_AI_NOTE) and "분류" not in text and "AI 초안:" not in text
    failed = format_alert(view, link, None, "(AI 초안 실패: 연결 오류)")
    assert failed.endswith("(AI 초안 실패: 연결 오류)")


def test_empty_draft_shows_placeholder():
    view, link = make()
    empty = DraftResult("기타", "사람에게", "불확실", "")
    assert format_alert(view, link, empty).endswith("(초안 없음)")


def test_long_fields_are_trimmed_with_ellipsis():
    view, link = make(question="가" * 500, name="나" * 60)
    text = format_alert(view, link, AI)
    assert "나" * 29 + "…" in text and "나" * 30 not in text
    assert '"' + "가" * 299 + '…"' in text


def test_option_is_shortened_and_missing_option_shows_quantity_only():
    link = link_customer_inquiry({"productOrderIdList": "2026100700001-1"}, [detail(option="옵션" * 30)])
    assert "옵션" * 9 + "옵…" in format_alert(make(link=link)[0], link, AI)
    bare = link_customer_inquiry({"productOrderIdList": "2026100700001-1"}, [detail(option="")])
    assert "· ×2" in format_alert(make(link=bare)[0], bare, AI)


def test_whole_message_stays_under_telegram_limit():
    view, link = make(question="문" * 500)
    huge = DraftResult("기타", "사람에게", "이" * 5000, "초" * 9000)
    text = format_alert(view, link, huge)
    assert len(text) <= TELEGRAM_LIMIT and text.endswith("…")
    assert len(format_alert(view, link, None, "x" * 9000)) <= TELEGRAM_LIMIT


def test_trim():
    assert trim("abc", 3) == "abc" and trim("abcd", 3) == "ab…" and trim(None, 3) == ""


@pytest.mark.parametrize("draft", [AI, None])
def test_no_personal_values_in_alert(draft):
    link = link_customer_inquiry({"productOrderIdList": "2026100700001-1"}, [detail()])
    view, _ = make(link=link)
    text = format_alert(view, link, draft)
    for secret in PERSONAL + ["2026100700001"]:
        assert secret not in text
