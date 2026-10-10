"""Inquiry API calls: paging, 365-day splitting and de-duplication — no network (fake session)."""
from datetime import date

import pytest

from tests.test_naver_client import FakeResponse, FakeSession
from worker.cs_collect import summarize_inquiries, summarize_qnas
from worker.inquiry_client import (
    INQUIRY_URL,
    QNA_URL,
    _spans,
    fetch_customer_inquiries,
    fetch_qnas,
)
from worker.naver_client import NaverApiError


def qna_page(ids, last, total_pages=2):
    return FakeResponse({"contents": [{"questionId": i, "answered": True} for i in ids],
                         "last": last, "totalPages": total_pages})


def inquiry_page(ids, last):
    return FakeResponse({"content": [{"inquiryNo": i} for i in ids], "last": last})


def test_qnas_follow_pages_until_last():
    session = FakeSession(qna_page([1, 2], last=False), qna_page([3], last=True))
    items = fetch_qnas("tok", date(2026, 9, 1), date(2026, 9, 30), session)
    assert [i["questionId"] for i in items] == [1, 2, 3]
    pages = [call[2]["params"]["page"] for call in session.calls]
    assert pages == [1, 2]
    params = session.calls[0][2]["params"]
    assert session.calls[0][1] == QNA_URL
    assert params["fromDate"] == "2026-09-01T00:00:00.000+09:00"
    assert params["toDate"] == "2026-09-30T23:59:59.000+09:00"
    assert session.calls[0][2]["headers"]["Authorization"] == "Bearer tok"


def test_qnas_stop_at_total_pages_even_without_last_flag():
    session = FakeSession(FakeResponse({"contents": [{"questionId": 1}], "totalPages": 1}))
    assert len(fetch_qnas("tok", date(2026, 9, 1), date(2026, 9, 2), session)) == 1
    assert len(session.calls) == 1


def test_qnas_drop_duplicates_that_shift_between_pages():
    session = FakeSession(qna_page([1, 2], last=False), qna_page([2, 3], last=True))
    items = fetch_qnas("tok", date(2026, 9, 1), date(2026, 9, 30), session)
    assert [i["questionId"] for i in items] == [1, 2, 3]


def test_spans_split_into_inclusive_365_day_chunks():
    spans = _spans(date(2025, 1, 1), date(2026, 1, 1), 365)
    assert spans == [(date(2025, 1, 1), date(2025, 12, 31)), (date(2026, 1, 1), date(2026, 1, 1))]
    assert _spans(date(2026, 1, 1), date(2026, 1, 1), 365) == [(date(2026, 1, 1), date(2026, 1, 1))]


def test_inquiries_over_a_year_are_split_into_two_calls():
    session = FakeSession(inquiry_page([1], last=True), inquiry_page([2], last=True))
    items = fetch_customer_inquiries("tok", date(2025, 1, 1), date(2026, 1, 1), session)
    assert [i["inquiryNo"] for i in items] == [1, 2]
    first, second = (call[2]["params"] for call in session.calls)
    assert session.calls[0][1] == INQUIRY_URL
    assert (first["startSearchDate"], first["endSearchDate"]) == ("2025-01-01", "2025-12-31")
    assert (second["startSearchDate"], second["endSearchDate"]) == ("2026-01-01", "2026-01-01")


def test_inquiries_follow_pages_within_a_span():
    session = FakeSession(inquiry_page([1], last=False), inquiry_page([2], last=True))
    items = fetch_customer_inquiries("tok", date(2026, 9, 1), date(2026, 9, 30), session)
    assert [i["inquiryNo"] for i in items] == [1, 2]
    assert [call[2]["params"]["page"] for call in session.calls] == [1, 2]


def test_refused_inquiry_call_raises():
    session = FakeSession(FakeResponse({"code": "4003", "message": "최대 365일만 조회 가능합니다."}, 400))
    with pytest.raises(NaverApiError, match="4003"):
        fetch_customer_inquiries("tok", date(2026, 9, 1), date(2026, 9, 30), session)


def test_summaries_count_without_question_text():
    qnas = [{"createDate": "2026-09-03T10:00:00", "productName": "오징어", "answered": True,
             "question": "비밀 질문"},
            {"createDate": "2026-10-01T10:00:00", "productName": "오징어", "answered": False}]
    inquiries = [{"inquiryRegistrationDateTime": "2026-09-05T10:00:00", "category": "배송",
                  "answered": True, "orderId": "1", "customerName": "홍길동", "inquiryContent": "비밀"}]
    text = "\n".join(summarize_qnas(qnas) + summarize_inquiries(inquiries))
    assert "상품문의 2건 (답변 완료 1건, 미답변 1건)" in text
    assert "'2026-09': 1" in text and "오징어" in text
    assert "고객문의 1건 (답변 완료 1건, 주문 연결 1건)" in text and "'배송': 1" in text
    assert "비밀" not in text and "홍길동" not in text


def test_rate_limited_call_waits_and_retries():
    waits = []
    session = FakeSession(FakeResponse({"code": "GW.RATE_LIMIT"}, 429), qna_page([1], last=True))
    items = fetch_qnas("tok", date(2026, 9, 1), date(2026, 9, 2), session, sleep=waits.append)
    assert [i["questionId"] for i in items] == [1]
    assert waits == [1.0]


def test_rate_limit_that_never_clears_raises():
    session = FakeSession(*[FakeResponse({"code": "GW.RATE_LIMIT", "message": "제한"}, 429)] * 5)
    with pytest.raises(NaverApiError, match="429"):
        fetch_customer_inquiries("tok", date(2026, 9, 1), date(2026, 9, 2), session, sleep=lambda _: None)
