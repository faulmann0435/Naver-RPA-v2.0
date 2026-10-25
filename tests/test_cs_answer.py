"""Posting an answer: request shape, safety switch, already-answered refusal, text checks — no network."""
from datetime import date, datetime

import pytest

from worker import cs_answer
from worker.cs_answer import (
    INQUIRY_ANSWER_URL,
    INQUIRY_BODY_FIELD,
    MAX_ANSWER_CHARS,
    QNA_ANSWER_URL,
    QNA_BODY_FIELD,
    AlreadyAnswered,
    CsAnswerError,
    is_answered,
    post_inquiry_answer,
    post_qna_answer,
    submit_answer,
    validate_answer,
)
from worker.naver_adapter import KST

NOW = datetime(2026, 10, 10, 12, 0, tzinfo=KST)
CREATED = "2026-10-09T09:00:00+09:00"


class FakeResponse:
    def __init__(self, body=None, status=200):
        self._body, self.status_code, self.text = body if body is not None else {}, status, str(body)

    def json(self):
        return self._body


class FakeSession:
    def __init__(self, *responses):
        self.responses, self.calls = list(responses), []

    def _answer(self, method, url, kwargs):
        self.calls.append((method, url, kwargs))
        return self.responses.pop(0) if self.responses else FakeResponse()

    def put(self, url, **kwargs):
        return self._answer("PUT", url, kwargs)

    def post(self, url, **kwargs):
        return self._answer("POST", url, kwargs)


# --- request shape ------------------------------------------------------------------------------

def test_qna_answer_is_a_put_with_comment_content():
    session = FakeSession()
    post_qna_answer("tok", "123", "답변입니다", session)
    method, url, kwargs = session.calls[0]
    assert (method, url) == ("PUT", QNA_ANSWER_URL.format(question_id="123"))
    assert url == "https://api.commerce.naver.com/external/v1/contents/qnas/123"
    assert kwargs["json"] == {QNA_BODY_FIELD: "답변입니다"} == {"commentContent": "답변입니다"}
    assert kwargs["headers"] == {"Authorization": "Bearer tok"}


def test_inquiry_answer_is_a_post_with_answer_comment():
    session = FakeSession()
    post_inquiry_answer("tok", "77", "답변입니다", session)
    method, url, kwargs = session.calls[0]
    assert (method, url) == ("POST", INQUIRY_ANSWER_URL.format(inquiry_no="77"))
    assert url.endswith("/v1/pay-merchant/inquiries/77/answer")
    assert kwargs["json"] == {INQUIRY_BODY_FIELD: "답변입니다"} == {"answerComment": "답변입니다"}


def test_rate_limit_is_retried(monkeypatch):
    monkeypatch.setattr("worker.order_link.time.sleep", lambda s: None)
    session = FakeSession(FakeResponse(status=429), FakeResponse())
    post_qna_answer("tok", "1", "x", session)
    assert len(session.calls) == 2


def test_client_error_shows_naver_code_and_message():
    session = FakeSession(FakeResponse({"code": "InvalidInput", "message": "commentContent is required"}, 400))
    with pytest.raises(CsAnswerError) as err:
        post_qna_answer("tok", "1", "x", session)
    text = str(err.value)
    assert "HTTP 400" in text and "InvalidInput" in text and "commentContent is required" in text
    assert "개발자" in text


def test_server_error_is_reported_without_field_hint():
    with pytest.raises(CsAnswerError, match="HTTP 500") as err:
        post_inquiry_answer("tok", "1", "x", FakeSession(FakeResponse({"message": "boom"}, 500)))
    assert "항목 이름" not in str(err.value)


# --- text checks --------------------------------------------------------------------------------

def test_validate_strips_and_refuses_empty_or_too_long():
    assert validate_answer("  안녕하세요  ") == "안녕하세요"
    for bad in ("", "   \n", None):
        with pytest.raises(CsAnswerError, match="비어"):
            validate_answer(bad)
    assert validate_answer("가" * MAX_ANSWER_CHARS)
    with pytest.raises(CsAnswerError, match="너무 깁니다"):
        validate_answer("가" * (MAX_ANSWER_CHARS + 1))


# --- already answered ---------------------------------------------------------------------------

def test_is_answered_checks_the_right_channel_by_id():
    qnas = lambda token, start, end: [{"questionId": 5, "answered": True}, {"questionId": 6, "answered": False}]
    inquiries = lambda token, start, end: [{"inquiryNo": 9, "answered": True}]
    args = (date(2026, 10, 9), date(2026, 10, 10), "tok")
    assert is_answered("상품문의", "5", *args, fetch_qna_list=qnas) is True
    assert is_answered("상품문의", "6", *args, fetch_qna_list=qnas) is False
    assert is_answered("상품문의", "404", *args, fetch_qna_list=qnas) is False
    assert is_answered("고객문의", "9", *args, fetch_inquiry_list=inquiries) is True


# --- submit_answer: switch, refusal -------------------------------------------------------------

class Spy:
    def __init__(self):
        self.calls = []

    def __call__(self, *args):
        self.calls.append(args)


def submit(channel="상품문의", enabled=True, answered=False, **kw):
    spies = {"qna": Spy(), "inq": Spy()}
    result = submit_answer(channel, "5", CREATED, "  답변  ", enabled=enabled, now=lambda: NOW, token="tok",
                           answered_check=lambda *a: answered, qna_poster=spies["qna"],
                           inquiry_poster=spies["inq"], **kw)
    return result, spies


def test_switch_on_posts_to_the_right_channel():
    result, spies = submit("상품문의")
    assert result.simulated is False and spies["qna"].calls == [("tok", "5", "답변", None)] and not spies["inq"].calls
    result, spies = submit("고객문의")
    assert spies["inq"].calls == [("tok", "5", "답변", None)] and not spies["qna"].calls


def test_switch_off_is_simulated_and_sends_nothing():
    result, spies = submit(enabled=False)
    assert result.simulated is True and not spies["qna"].calls and not spies["inq"].calls


def test_already_answered_is_refused_even_in_test_mode():
    for enabled in (True, False):
        with pytest.raises(AlreadyAnswered, match="이미 판매자센터에서 답변된 문의입니다"):
            submit(enabled=enabled, answered=True)


def test_bad_text_is_refused_before_any_call():
    with pytest.raises(CsAnswerError):
        submit_answer("상품문의", "5", CREATED, " ", enabled=True, token="tok",
                      answered_check=lambda *a: pytest.fail("no fetch"))


def test_answered_check_gets_creation_day_and_today():
    seen = []
    submit_answer("상품문의", "5", CREATED, "x", enabled=False, now=lambda: NOW, token="tok",
                  answered_check=lambda *a: seen.append(a) or False)
    assert seen == [("상품문의", "5", date(2026, 10, 9), date(2026, 10, 10), "tok")]


def test_switch_reads_env_then_dotenv_on_every_call(monkeypatch, tmp_path):
    env = tmp_path / ".env"
    monkeypatch.setattr(cs_answer, "ENV_FILE", env)
    monkeypatch.delenv("CS_ANSWER_ENABLED", raising=False)
    assert cs_answer.answer_enabled() is False
    env.write_text("CS_ANSWER_ENABLED=1\n", encoding="utf-8")
    assert cs_answer.answer_enabled() is True
    monkeypatch.setenv("CS_ANSWER_ENABLED", "0")
    assert cs_answer.answer_enabled() is False
