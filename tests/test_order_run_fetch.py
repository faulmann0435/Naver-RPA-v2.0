"""fetch_paid_items: the bot's API calls must survive Naver's short rate-limit refusals (GW.RATE_LIMIT)."""
import json

from tests.test_naver_client import CREDS, FakeResponse, FakeSession
from worker import order_run


def _limited():
    return FakeResponse({"code": "GW.RATE_LIMIT", "message": "요청이 많아 서비스를 일시적으로 사용할 수 없습니다."}, 429)


def test_rate_limited_calls_are_retried_not_reported(monkeypatch):
    monkeypatch.setattr(order_run, "load_credentials", lambda: CREDS)
    monkeypatch.setattr(order_run, "RETRY_SLEEP", lambda _: None)
    session = FakeSession(
        FakeResponse({"access_token": "t"}),
        _limited(),
        FakeResponse({"data": {"lastChangeStatuses": [{"productOrderId": 11}]}}),
        _limited(),
        FakeResponse({"data": [{"productOrder": {"productOrderId": "11"}}]}),
    )
    items = order_run.fetch_paid_items(days=1, session=session)
    assert items == [{"productOrder": {"productOrderId": "11"}}]
    assert not session.responses


def test_fetch_asks_for_every_change_not_only_payments(monkeypatch):
    """A paid order whose LAST change is not the payment (address changed afterwards) must still come back."""
    monkeypatch.setattr(order_run, "load_credentials", lambda: CREDS)
    session = FakeSession(
        FakeResponse({"access_token": "t"}),
        FakeResponse({"data": {"lastChangeStatuses": [
            {"productOrderId": 11, "lastChangedType": "DELIVERY_ADDRESS_CHANGED"}]}}),
        FakeResponse({"data": [{"productOrder": {"productOrderId": "11", "productOrderStatus": "PAYED"}}]}),
    )
    items = order_run.fetch_paid_items(days=1, session=session)
    assert [i["productOrder"]["productOrderId"] for i in items] == ["11"]
    _, _, query = session.calls[1]
    assert query["params"].get("lastChangedType") is None


def test_no_changes_means_no_detail_call(monkeypatch):
    monkeypatch.setattr(order_run, "load_credentials", lambda: CREDS)
    session = FakeSession(FakeResponse({"access_token": "t"}), FakeResponse({"data": {"lastChangeStatuses": []}}))
    assert order_run.fetch_paid_items(days=1, session=session) == []
    assert not session.responses


def test_shipped_and_cancelled_orders_are_not_counted_as_claims(tmp_path, monkeypatch):
    def po(i, status, claim=None):
        return {"order": {"orderId": "9"}, "productOrder": {
            "productOrderId": i, "productOrderStatus": status, "claimType": claim, "productName": "초벌장어",
            "productOption": "b", "quantity": 1, "packageNumber": "7", "productId": "5",
            "shippingAddress": {"zipCode": "12345", "name": "홍길동", "baseAddress": "서울시 중구"}}}

    monkeypatch.setattr(order_run, "process_orders",
                        lambda *a, **k: type("R", (), {"files": [], "stats": {}, "unmatched": []})())
    engine = lambda: type("E", (), {"config": None, "dictionary": None, "settings": None, "source": "t"})()
    items = [po("1", "PAYED"), po("2", "PAYED", "CANCEL"), po("3", "DELIVERING"), po("4", "CANCELED", "CANCEL")]
    run = order_run.run_orders(fetch=lambda: items, engine_loader=engine, ledger=tmp_path / "l.json")
    assert run.product_order_ids == ("1",) and run.left_out_claims == 1


def _po(i, place="OK"):
    return {"order": {"orderId": "9"}, "productOrder": {
        "productOrderId": i, "productOrderStatus": "PAYED", "placeOrderStatus": place, "productName": "초벌장어",
        "productOption": "b", "quantity": 1, "packageNumber": "7", "productId": "5",
        "shippingAddress": {"zipCode": "12345", "name": "홍길동", "baseAddress": "서울시 중구"}}}


def test_naver_counts_split_like_the_seller_center_and_include_shared_orders(tmp_path, monkeypatch):
    monkeypatch.setattr(order_run, "process_orders",
                        lambda *a, **k: type("R", (), {"files": [], "stats": {}, "unmatched": []})())
    engine = lambda: type("E", (), {"config": None, "dictionary": None, "settings": None, "source": "t"})()
    ledger = tmp_path / "l.json"
    order_run.mark_shared(("1",), ledger)
    items = [_po("1"), _po("2"), _po("3", "NOT_YET")]
    run = order_run.run_orders(fetch=lambda: items, engine_loader=engine, ledger=ledger)
    assert (run.paid, run.new_orders, run.waiting_orders) == (2, 1, 2)  # ledger order 1 still counted
    text = order_run.summary(run)
    assert "네이버 발송 전 주문 3건 = 신규주문 1 + 발송대기 2" in text and "이번 발주서에 넣은 주문 2건" in text


def test_order_without_address_waits_and_never_enters_the_ledger(tmp_path, monkeypatch):
    # e.g. 선물하기 수락대기: paid, but no address yet -> it used to be a blank line on the 발주서
    monkeypatch.setattr(order_run, "process_orders",
                        lambda *a, **k: type("R", (), {"files": [], "stats": {}, "unmatched": []})())
    engine = lambda: type("E", (), {"config": None, "dictionary": None, "settings": None, "source": "t"})()
    gift = {**_po("2", "NOT_YET"), "order": {"orderId": "9", "ordererName": "김선물", "ordererTel": "010-1"}}
    gift["productOrder"] = {**gift["productOrder"], "shippingAddress": {}, "giftReceivingStatus": "WAIT"}
    log = tmp_path / "status.json"
    for include_shared in (False, True):
        run = order_run.run_orders(fetch=lambda: [_po("1"), gift], engine_loader=engine,
                                   ledger=tmp_path / "l.json", include_shared=include_shared, address_log=log)
        assert run.product_order_ids == ("1",) and run.paid == 1
        assert run.address_waiting_names == ("김선물(초벌장어)",)
        assert (run.new_orders, run.waiting_orders) == (0, 1)  # left out of Naver counts like the Excel export
    text = order_run.summary(run)
    assert "🎁 배송지 입력 대기(선물 수락대기 등) 1건 제외: 김선물(초벌장어)" in text


def test_address_waiting_log_keeps_status_fields_but_no_personal_data(tmp_path):
    item = {"order": {"orderId": "9", "ordererName": "김선물", "ordererTel": "010-1", "paymentDate": "d"},
            "productOrder": {"productOrderId": "2", "productOrderStatus": "PAYED", "giftReceivingStatus": "WAIT",
                             "shippingAddress": {}, "productName": "장어"}}
    log = tmp_path / "status.json"
    order_run.log_address_waiting([item], log)
    logged = json.loads(log.read_text(encoding="utf-8"))["2"]
    assert logged["productOrder"] == {"productOrderStatus": "PAYED", "giftReceivingStatus": "WAIT"}
    assert logged["order"] == {"paymentDate": "d"}
    order_run.log_address_waiting([], tmp_path / "none.json")
    assert not (tmp_path / "none.json").exists()


def test_address_waiting_line_also_shows_when_nothing_is_new():
    text = order_run.summary(_result(paid=0, product_order_ids=(), address_waiting_names=("김선물(장어)",)))
    assert text.startswith("📭") and "🎁 배송지 입력 대기(선물 수락대기 등) 1건 제외" in text
    assert "배송지 입력 대기" not in order_run.summary(_result())


def _result(**kw):
    base = dict(files=[], product_order_ids=("1",), paid=1, left_out_claims=0, already_shared=0,
                unmatched=0, needs_review=0, source="t", new_orders=0, waiting_orders=1)
    return order_run.RunResult(**{**base, **kw})


def test_unregistered_items_warn_but_review_marks_are_only_a_note():
    review_only = order_run.summary(_result(needs_review=88))
    assert "✅ 미등록 품목 없음" in review_only and "⚠️" not in review_only
    assert "사전 검토 표시가 남은 품목 88줄 (발주서에는 정상 반영" in review_only
    both = order_run.summary(_result(unmatched=4, needs_review=88))
    assert "⚠️ 미등록 품목 4건 → v15 앱에서 등록한 뒤" in both and "88줄" in both


def test_nothing_new_still_shows_the_naver_counts():
    text = order_run.summary(_result(paid=0, product_order_ids=(), already_shared=1, new_orders=0, waiting_orders=1))
    assert text.startswith("📭 새로 발주할 주문이 없습니다") and "네이버 발송 전 주문 1건" in text
