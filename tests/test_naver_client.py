"""Naver API client: signature, token, paging and chunking — no network (fake session)."""
import base64
from datetime import datetime, timedelta

import bcrypt
import pytest

from worker.naver_adapter import KST
from worker.naver_client import (
    Credentials,
    NaverApiError,
    fetch_changed_ids,
    fetch_details,
    fetch_token,
    load_credentials,
    make_signature,
)

CREDS = Credentials("client-id", bcrypt.gensalt(rounds=4).decode())


class FakeResponse:
    def __init__(self, body, status=200):
        self._body, self.status_code, self.ok, self.text = body, status, status < 400, str(body)

    def json(self):
        return self._body


class FakeSession:
    def __init__(self, *responses):
        self.responses = list(responses)
        self.calls = []

    def _next(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        return self.responses.pop(0)

    def get(self, url, **kwargs):
        return self._next("GET", url, **kwargs)

    def post(self, url, **kwargs):
        return self._next("POST", url, **kwargs)


def test_signature_is_bcrypt_of_id_and_timestamp_with_secret_as_salt():
    sign = make_signature(CREDS, 1700000000000)
    hashed = base64.b64decode(sign)
    assert bcrypt.checkpw(b"client-id_1700000000000", hashed)


def test_token_request_uses_self_type_and_returns_token():
    session = FakeSession(FakeResponse({"access_token": "tok"}))
    assert fetch_token(CREDS, session) == "tok"
    data = session.calls[0][2]["data"]
    assert data["type"] == "SELF" and data["grant_type"] == "client_credentials"


def test_refused_call_raises_with_naver_message():
    session = FakeSession(FakeResponse({"code": "GW.IP_NOT_ALLOWED", "message": "허용되지 않은 IP"}, 403))
    with pytest.raises(NaverApiError, match="IP_NOT_ALLOWED"):
        fetch_token(CREDS, session)


def test_changed_ids_follow_more_pages_and_drop_duplicates():
    page1 = {"data": {"lastChangeStatuses": [{"productOrderId": "1"}, {"productOrderId": "2"}],
                      "more": {"moreFrom": "2026-10-10T10:00:00.000+09:00", "moreSequence": "s1"}}}
    page2 = {"data": {"lastChangeStatuses": [{"productOrderId": "2"}, {"productOrderId": "3"}]}}
    session = FakeSession(FakeResponse(page1), FakeResponse(page2))
    end = datetime(2026, 10, 10, 12, tzinfo=KST)
    assert fetch_changed_ids("tok", end - timedelta(hours=5), end, session=session) == ["1", "2", "3"]
    assert session.calls[1][2]["params"]["moreSequence"] == "s1"


def test_changed_ids_split_ranges_longer_than_a_day():
    empty = FakeResponse({"data": {"lastChangeStatuses": []}})
    session = FakeSession(empty, empty, empty)
    end = datetime(2026, 10, 10, 12, tzinfo=KST)
    fetch_changed_ids("tok", end - timedelta(days=3), end, session=session)
    assert len(session.calls) == 3


def test_details_are_requested_300_at_a_time():
    session = FakeSession(FakeResponse({"data": [{"n": 1}]}), FakeResponse({"data": [{"n": 2}]}))
    items = fetch_details("tok", [str(i) for i in range(301)], session=session)
    assert items == [{"n": 1}, {"n": 2}]
    assert [len(c[2]["json"]["productOrderIds"]) for c in session.calls] == [300, 1]


def test_missing_keys_fail_clearly(tmp_path, monkeypatch):
    monkeypatch.delenv("NAVER_COMMERCE_CLIENT_ID", raising=False)
    monkeypatch.delenv("NAVER_COMMERCE_CLIENT_SECRET", raising=False)
    with pytest.raises(NaverApiError, match="설정되지 않았습니다"):
        load_credentials(tmp_path / ".env")


def test_keys_are_read_from_env_file(tmp_path, monkeypatch):
    monkeypatch.delenv("NAVER_COMMERCE_CLIENT_ID", raising=False)
    monkeypatch.delenv("NAVER_COMMERCE_CLIENT_SECRET", raising=False)
    env = tmp_path / ".env"
    env.write_text("NAVER_COMMERCE_CLIENT_ID=a\nNAVER_COMMERCE_CLIENT_SECRET=$2a$04$x\n", encoding="utf-8")
    assert load_credentials(env) == Credentials("a", "$2a$04$x")
