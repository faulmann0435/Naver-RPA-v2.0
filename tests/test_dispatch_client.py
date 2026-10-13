"""Dispatch client: chunking, body shape, result merging, validation, safety switch — no network."""
from datetime import datetime, timedelta

import pytest

from worker.dispatch_client import (
    DISPATCH_URL,
    DispatchItem,
    dispatch_enabled,
    dispatch_or_simulate,
    send_dispatch,
)
from worker.naver_adapter import KST
from worker.naver_client import NaverApiError

NOW = datetime(2026, 10, 13, 12, 0, tzinfo=KST)


class FakeResponse:
    def __init__(self, body, status=200):
        self._body, self.status_code, self.ok, self.text = body, status, status < 400, str(body)

    def json(self):
        return self._body


class FakeSession:
    def __init__(self, *responses):
        self.responses = list(responses)
        self.calls = []

    def post(self, url, **kwargs):
        self.calls.append((url, kwargs))
        reply = self.responses.pop(0)
        if isinstance(reply, Exception):
            raise reply
        return reply


def _items(n):
    return [DispatchItem(str(i), f"T{i}") for i in range(n)]


def _ok(items):
    return FakeResponse({"data": {"successProductOrderIds": [i.product_order_id for i in items],
                                  "failProductOrderInfos": []}})


def _send(items, session, **kw):
    return send_dispatch(items, session=session, token="tok", now=lambda: NOW, **kw)


def test_65_items_go_out_in_30_30_5():
    items = _items(65)
    session = FakeSession(_ok(items[:30]), _ok(items[30:60]), _ok(items[60:]))
    result = _send(items, session)
    sizes = [len(c[1]["json"]["dispatchProductOrders"]) for c in session.calls]
    assert sizes == [30, 30, 5] and len(result.success_ids) == 65 and not result.failures


def test_request_shape():
    session = FakeSession(_ok(_items(1)))
    _send(_items(1), session)
    url, kwargs = session.calls[0]
    assert url == DISPATCH_URL and kwargs["headers"] == {"Authorization": "Bearer tok"}
    assert kwargs["json"] == {"dispatchProductOrders": [{
        "productOrderId": "0", "deliveryMethod": "DELIVERY", "deliveryCompanyCode": "HANJIN",
        "trackingNumber": "T0", "dispatchDate": "2026-10-13T12:00:00.000+09:00"}]}


def test_success_and_failures_are_merged():
    items = _items(3)
    reply = FakeResponse({"data": {"successProductOrderIds": ["0", "1"], "failProductOrderInfos": [
        {"productOrderId": "2", "code": "E1", "message": "이미 발송됨"}]}})
    result = _send(items, FakeSession(reply))
    assert result.success_ids == ("0", "1")
    assert [(f.id, f.code, f.message) for f in result.failures] == [("2", "E1", "이미 발송됨")]
    assert result.simulated is False


def test_ids_missing_from_the_answer_count_as_failures():
    result = _send(_items(2), FakeSession(FakeResponse({"data": {"successProductOrderIds": ["0"]}})))
    assert [f.id for f in result.failures] == ["1"]


def test_error_after_first_chunk_keeps_what_succeeded():
    items = _items(65)
    session = FakeSession(_ok(items[:30]), FakeResponse({"code": "X", "message": "boom"}, 500))
    result = _send(items, session)
    assert len(session.calls) == 2 and len(result.success_ids) == 30
    assert [f.id for f in result.failures] == [str(i) for i in range(30, 65)]
    assert {f.code for f in result.failures} == {"REQUEST_FAILED"} and "boom" in result.failures[0].message


def test_network_exception_is_reported_not_raised():
    result = _send(_items(2), FakeSession(NaverApiError("HTTP 403: GW.IP_NOT_ALLOWED")))
    assert result.success_ids == () and "IP_NOT_ALLOWED" in result.failures[0].message


@pytest.mark.parametrize("tracking", ["", "12 34", "12_34", "가나다", "1" * 51])
def test_bad_tracking_number_is_refused_before_any_call(tracking):
    session = FakeSession()
    with pytest.raises(ValueError, match="송장번호"):
        _send([DispatchItem("1", tracking)], session)
    assert session.calls == []


def test_dispatch_date_must_be_recent_and_not_future():
    session = FakeSession()
    with pytest.raises(ValueError, match="미래"):
        _send(_items(1), session, when=NOW + timedelta(minutes=1))
    with pytest.raises(ValueError, match="30일"):
        _send(_items(1), session, when=NOW - timedelta(days=31))
    assert session.calls == []


def test_switch_off_simulates_without_http():
    session = FakeSession()
    result = dispatch_or_simulate(_items(3), enabled=False, session=session, now=lambda: NOW)
    assert result.simulated and result.success_ids == ("0", "1", "2") and session.calls == []


def test_switch_on_sends_for_real():
    items = _items(2)
    session = FakeSession(_ok(items))
    result = dispatch_or_simulate(items, enabled=True, session=session, token="tok", now=lambda: NOW)
    assert not result.simulated and len(session.calls) == 1


def test_switch_reads_environment(monkeypatch):
    monkeypatch.setenv("INVOICE_DISPATCH_ENABLED", "1")
    assert dispatch_enabled()
    monkeypatch.setenv("INVOICE_DISPATCH_ENABLED", "0")
    assert not dispatch_enabled()


def test_each_item_carries_its_own_courier_code():
    items = [DispatchItem("1", "600663392321", "CJGLS"), DispatchItem("2", "45287820340", "KGB")]
    session = FakeSession(_ok(items))
    _send(items, session)
    codes = [o["deliveryCompanyCode"] for o in session.calls[0][1]["json"]["dispatchProductOrders"]]
    assert codes == ["CJGLS", "KGB"]
