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
