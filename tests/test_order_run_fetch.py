"""fetch_paid_items: the bot's API calls must survive Naver's short rate-limit refusals (GW.RATE_LIMIT)."""
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
            "shippingAddress": {"zipCode": "12345", "name": "홍길동"}}}

    monkeypatch.setattr(order_run, "process_orders",
                        lambda *a, **k: type("R", (), {"files": [], "stats": {}, "unmatched": []})())
    engine = lambda: type("E", (), {"config": None, "dictionary": None, "settings": None, "source": "t"})()
    items = [po("1", "PAYED"), po("2", "PAYED", "CANCEL"), po("3", "DELIVERING"), po("4", "CANCELED", "CANCEL")]
    run = order_run.run_orders(fetch=lambda: items, engine_loader=engine, ledger=tmp_path / "l.json")
    assert run.product_order_ids == ("1",) and run.left_out_claims == 1
