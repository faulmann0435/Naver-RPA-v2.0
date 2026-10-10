"""Inquiry -> order linking: summaries, matching rules, retry session — no network (fake sessions)."""
import json
from dataclasses import asdict
from datetime import datetime

import pytest

from tests.test_naver_client import FakeResponse, FakeSession
from worker import cs_link
from worker.naver_adapter import KST
from worker.naver_client import NaverApiError
from worker.order_link import (
    LinkResult,
    RetryingSession,
    clean_memo,
    fetch_recent_details,
    link_customer_inquiry,
    link_qna,
    missing_inquiry_ids,
    parse_time,
    summarize,
    to_prompt_text,
)

PERSONAL = ["홍길동", "010-1234-5678", "서울시 강남구 비밀로 1", "gotu****", "TRK123456"]


def detail(order_id="2026100700001", po_id="2026100700001-1", orderer="gotu****", product_id="555",
           paid="2026-10-07T12:42:00+09:00", status="PURCHASE_DECIDED", name="국산 풍천 장어", option="1kg",
           qty=2, delivery=True, memo="문앞", arrival=None):
    item = {"order": {"orderId": order_id, "paymentDate": paid, "ordererId": orderer,
                      "ordererName": "홍길동", "ordererTel": "010-1234-5678"},
            "productOrder": {"productOrderId": po_id, "productId": product_id, "productName": name,
                             "productOption": option, "quantity": qty, "productOrderStatus": status,
                             "shippingDueDate": "2026-10-08T23:59:59+09:00", "arrivalGuaranteeDate": arrival,
                             "expectedDeliveryCompany": "CJGLS", "shippingMemo": memo,
                             "shippingAddress": {"baseAddress": "서울시 강남구 비밀로 1"}}}
    if delivery:
        item["delivery"] = {"deliveryCompany": "HANJIN", "sendDate": "2026-10-07T15:10:00+09:00",
                            "deliveredDate": "2026-10-09T10:00:00+09:00", "trackingNumber": "TRK123456"}
    return item


def qna(masked="gotu****", product_id=555, created="2026-10-08T09:00:00+09:00"):
    return {"questionId": 1, "createDate": created, "productId": product_id, "maskedWriterId": masked,
            "question": "언제 와요?"}


# --- summarize -------------------------------------------------------------------------------

def test_summarize_groups_product_orders_by_order():
    items = [detail(po_id="a", name="장어", status="DELIVERED"), detail(po_id="b", name="오징어", status="PAYED"),
             detail(order_id="999", po_id="c")]
    first, second = summarize(items)
    assert first.product_order_ids == ("a", "b") and first.product_name == "장어 / 오징어"
    assert first.status == "DELIVERED / PAYED" and first.status_label == "배송완료 / 결제완료(발송 전)"
    assert first.quantity == 4 and second.order_id == "999"


def test_summarize_fields_and_courier_fallback():
    one = summarize([detail()])[0]
    assert one.status_label == "구매확정" and one.courier == "HANJIN" and one.memo == "문앞"
    assert one.paid_at == datetime(2026, 10, 7, 12, 42, tzinfo=KST)
    assert one.sent_at.hour == 15 and one.delivered_at.day == 9 and one.arrival_guarantee is None
    unsent = summarize([detail(delivery=False, status="PAYED")])[0]
    assert unsent.courier == "CJGLS" and unsent.sent_at is None


def test_unknown_status_shows_raw_code_and_missing_order_id_is_skipped():
    assert summarize([detail(status="WEIRD")])[0].status_label == "WEIRD"
    assert summarize([{"productOrder": {}}]) == []


def test_summary_has_no_personal_values():
    text = json.dumps([asdict(s) for s in summarize([detail()])], ensure_ascii=False, default=str)
    for secret in PERSONAL:
        assert secret not in text


def test_parse_time_handles_empty_naive_and_garbage():
    assert parse_time(None) is None and parse_time("nonsense") is None
    assert parse_time("2026-10-07T12:42:00").tzinfo == KST
    assert parse_time("2026-10-07T03:42:00Z") == datetime(2026, 10, 7, 12, 42, tzinfo=KST)


# --- customer inquiry ------------------------------------------------------------------------

def test_inquiry_exact_by_product_order_id_list():
    details = [detail(po_id="a"), detail(order_id="2", po_id="b")]
    result = link_customer_inquiry({"productOrderIdList": "x, b", "orderId": "1"}, details)
    assert result.kind == "정확" and [o.order_id for o in result.orders] == ["2"]


def test_inquiry_falls_back_to_order_id():
    result = link_customer_inquiry({"productOrderIdList": "", "orderId": "2026100700001"}, [detail()])
    assert result.kind == "정확"


def test_inquiry_without_match_is_none():
    result = link_customer_inquiry({"productOrderIdList": "zzz", "orderId": "nope"}, [detail()])
    assert result == LinkResult("없음", (), "주문 정보를 찾지 못했습니다")
    assert link_customer_inquiry({}, [detail()]).kind == "없음"


# --- Q&A -------------------------------------------------------------------------------------

def test_qna_single_match_is_a_guess():
    result = link_qna(qna(), [detail()])
    assert result.kind == "추정" and len(result.orders) == 1
    assert result.note == "작성자 아이디 앞부분·상품·결제 시각으로 추정한 주문입니다"


def test_qna_several_orders_keep_three_most_recent():
    details = [detail(order_id=str(n), po_id=f"p{n}", paid=f"2026-10-0{n}T10:00:00+09:00") for n in range(1, 6)]
    result = link_qna(qna(created="2026-10-09T09:00:00+09:00"), details)
    assert result.kind == "후보 여러 건"
    assert [o.order_id for o in result.orders] == ["5", "4", "3"]


def test_qna_two_product_orders_of_one_order_is_one_guess():
    result = link_qna(qna(), [detail(po_id="a"), detail(po_id="b")])
    assert result.kind == "추정" and result.orders[0].product_order_ids == ("a", "b")


@pytest.mark.parametrize("masked, details", [
    ("gotu*****", [detail()]),                                    # different mask length
    ("gotu****", [detail(product_id="999")]),                     # different product
    ("gotu****", [detail(paid="2026-10-08T10:00:00+09:00")]),     # paid after the question
    ("gotx****", [detail()]),                                     # different visible prefix
    ("g*******", [detail(orderer="g*******")]),                   # prefix too short
    ("", [detail()]),
    ("gotu****", [detail(paid=None)]),                            # no payment time
])
def test_qna_non_matches_are_none(masked, details):
    assert link_qna(qna(masked=masked), details).kind == "없음"


def test_qna_with_unreadable_question_time_is_none():
    assert link_qna(qna(created=None), [detail()]).kind == "없음"


# --- prompt text -----------------------------------------------------------------------------

def test_prompt_text_for_none():
    text = to_prompt_text(link_qna(qna(masked="zz****"), [detail()]))
    assert text == "주문 연결: 없음 (주문 정보를 모르므로 출고일·도착일을 단정하지 말 것)"


def test_prompt_text_for_guess_has_caution_and_no_personal_data():
    text = to_prompt_text(link_qna(qna(), [detail()]))
    lines = text.splitlines()
    assert lines[0] == "주문 연결: 추정"
    assert "결제 2026-10-07 12:42 / 구매확정 / 국산 풍천 장어 (1kg) ×2" in lines[1]
    assert "발송 2026-10-07 15:10 HANJIN" in lines[1] and "배송메모: 문앞" in lines[1]
    assert lines[-1] == "(추정 연결 — 날짜를 단정하지 말고 사장님 확인용으로만 사용)"
    for secret in PERSONAL:
        assert secret not in text


def test_prompt_text_for_exact_unsent_order_and_arrival_guarantee():
    item = detail(delivery=False, status="PAYED", option="", memo="", arrival="2026-10-12T00:00:00+09:00")
    text = to_prompt_text(link_customer_inquiry({"productOrderIdList": "2026100700001-1"}, [item]))
    assert text.startswith("주문 연결: 정확") and "추정 연결" not in text
    assert "결제완료(발송 전)" in text and "택배사 CJGLS" in text and "도착보장 2026-10-12 00:00" in text
    assert "()" not in text and "배송메모" not in text


def test_prompt_text_for_many_candidates():
    details = [detail(order_id=str(n), po_id=f"p{n}") for n in (1, 2)]
    text = to_prompt_text(link_qna(qna(), details))
    assert text.startswith("주문 연결: 후보 여러 건") and "단정하지 말고" in text.splitlines()[-1]


# --- retry session and fetching --------------------------------------------------------------

def limited():
    return FakeResponse({"code": "GW.RATE_LIMIT"}, 429)


def test_retrying_session_retries_get_and_post_on_429():
    waits = []
    session = FakeSession(limited(), limited(), FakeResponse({"ok": 1}), limited(), FakeResponse({"ok": 2}))
    http = RetryingSession(session, sleep=waits.append)
    assert http.get("u", params={}).json() == {"ok": 1}
    assert http.post("u", json={}).json() == {"ok": 2}
    assert waits == [1.0, 2.0, 1.0]


def test_retrying_session_gives_up_and_returns_last_429():
    waits = []
    session = FakeSession(*[limited() for _ in range(5)])
    resp = RetryingSession(session, sleep=waits.append).get("u")
    assert resp.status_code == 429 and waits == [1.0, 2.0, 4.0, 8.0]


def test_retrying_session_does_not_touch_other_errors():
    session = FakeSession(FakeResponse({}, 500))
    assert RetryingSession(session, sleep=lambda _: 1 / 0).get("u").status_code == 500


def test_fetch_recent_details_sends_no_changed_type_and_retries(monkeypatch):
    ids_page = {"data": {"lastChangeStatuses": [{"productOrderId": "1"}]}}
    empty = FakeResponse({"data": {"lastChangeStatuses": []}})
    session = FakeSession(limited(), FakeResponse(ids_page), empty,
                          FakeResponse({"data": [detail()]}))
    now = datetime(2026, 10, 10, 12, tzinfo=KST)
    real = RetryingSession
    monkeypatch.setattr("worker.order_link.RetryingSession", lambda inner: real(inner, sleep=lambda _: None))
    items = fetch_recent_details("tok", days=2, session=session, now=now)
    assert items == [detail()]
    gets = [c for c in session.calls if c[0] == "GET"]
    assert gets and all(c[2]["params"].get("lastChangedType") is None for c in gets)  # None: requests drops it
    assert session.calls[-1][0] == "POST"


def test_fetch_recent_details_skips_detail_call_when_nothing_changed(monkeypatch):
    monkeypatch.setattr("worker.order_link.fetch_changed_ids", lambda *a, **k: [])
    monkeypatch.setattr("worker.order_link.fetch_details", lambda *a, **k: pytest.fail("no call expected"))
    assert fetch_recent_details("tok", session=FakeSession()) == []


# --- cs_link CLI helpers ---------------------------------------------------------------------

def test_console_line_is_masked_and_records_keep_question(tmp_path):
    inquiry = {"inquiryNo": 7, "orderId": "2026100700001", "productOrderIdList": "2026100700001-1",
               "inquiryRegistrationDateTime": "2026-10-09T08:00:00+09:00", "answered": False,
               "inquiryContent": "비밀 내용"}
    records = cs_link.build_records([qna()], [inquiry], [detail()])
    lines = [cs_link.console_line(r) for r in records]
    assert lines[0].startswith("[상품문의] 2026-10-08 09:00 답변 아니오 | 연결: 추정 | 주문 …0001 결제 2026-10-07 12:42")
    assert lines[1].startswith("[고객문의]") and "연결: 정확" in lines[1]
    assert not any("언제 와요" in line or "비밀" in line or "2026100700001" in line for line in lines)
    path = cs_link.save(tmp_path / "out", records)
    saved = json.loads(path.read_text(encoding="utf-8"))
    assert saved[0]["question"] == "언제 와요?" and saved[0]["orders"][0]["paid_at"].startswith("2026-10-07T12:42")
    assert "비밀 내용" in path.read_text(encoding="utf-8")


def test_console_line_without_order():
    record = cs_link.build_records([qna(masked="zz****")], [], [])[0]
    assert cs_link.console_line(record).endswith("연결: 없음")


def test_main_prints_korean_message_and_returns_1_on_api_error(monkeypatch, capsys):
    def boom():
        raise NaverApiError("키 없음")
    monkeypatch.setattr(cs_link, "load_credentials", boom)
    assert cs_link.main(["--days", "1"]) == 1
    assert "네이버 API 오류: 키 없음" in capsys.readouterr().out


def test_main_runs_end_to_end_with_stubs(monkeypatch, tmp_path, capsys):
    monkeypatch.setattr(cs_link, "load_credentials", lambda: None)
    monkeypatch.setattr(cs_link, "fetch_token", lambda creds: "tok")
    monkeypatch.setattr(cs_link, "fetch_qnas", lambda *a, **k: [qna()])
    monkeypatch.setattr(cs_link, "fetch_customer_inquiries", lambda *a, **k: [])
    monkeypatch.setattr(cs_link, "fetch_recent_details", lambda *a, **k: [detail()])
    assert cs_link.main(["--days", "3", "--lookback", "5", "--out", str(tmp_path)]) == 0
    assert (tmp_path / "주문연결.json").exists()
    assert "연결: 추정" in capsys.readouterr().out


# --- memo masking and inquiries older than the lookback ----------------------------------------

@pytest.mark.parametrize("memo, expected", [
    ("부재시 010-1234-5678 연락", "부재시 [숫자] 연락"),
    ("공동현관 #1234* 문앞", "공동현관 [숫자] 문앞"),
    ("1kg씩 따로 포장", "1kg씩 따로 포장"),
    (None, ""),
])
def test_clean_memo_hides_phones_and_door_codes(memo, expected):
    assert clean_memo(memo) == expected


def test_summary_memo_is_cleaned():
    [summary] = summarize([detail(memo="경비실 맡겨주세요 01098765432")])
    assert summary.memo == "경비실 맡겨주세요 [숫자]"


def test_missing_inquiry_ids_lists_only_unknown_product_orders():
    inquiries = [{"productOrderIdList": "A-1, B-1"}, {"productOrderIdList": ""}, {}]
    assert missing_inquiry_ids(inquiries, [detail(po_id="A-1")]) == ["B-1"]


def test_main_fetches_inquiry_orders_older_than_the_lookback(monkeypatch, tmp_path, capsys):
    asked = []
    monkeypatch.setattr(cs_link, "load_credentials", lambda: None)
    monkeypatch.setattr(cs_link, "fetch_token", lambda creds: "tok")
    monkeypatch.setattr(cs_link, "fetch_qnas", lambda *a, **k: [])
    monkeypatch.setattr(cs_link, "fetch_customer_inquiries",
                        lambda *a, **k: [{"inquiryNo": 7, "productOrderIdList": "OLD-1", "answered": False}])
    monkeypatch.setattr(cs_link, "fetch_recent_details", lambda *a, **k: [])

    def fake_details(token, ids, session):
        asked.extend(ids)
        return [detail(order_id="OLD", po_id="OLD-1")]
    monkeypatch.setattr(cs_link, "fetch_details", fake_details)
    assert cs_link.main(["--out", str(tmp_path)]) == 0
    assert asked == ["OLD-1"]
    assert "연결: 정확" in capsys.readouterr().out
