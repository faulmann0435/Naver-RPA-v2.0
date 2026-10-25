"""Post an answer to a customer question on Naver: the CS feature's only WRITE to the store.

    product Q&A   PUT  /v1/contents/qnas/{questionId}                    {"commentContent": text}
    1:1 inquiry   POST /v1/pay-merchant/inquiries/{inquiryNo}/answer     {"answerComment": text}

Safety switch: nothing is really posted unless CS_ANSWER_ENABLED=1 is set in the environment or
.env (read on every call). Without it submit_answer() only pretends (simulated=True) and makes no
write call. Before posting, the question is fetched again and refused when it is already answered
(the owner may have answered in the seller center meanwhile).
"""
from __future__ import annotations

import os
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime

import requests

from worker.inquiry_client import fetch_customer_inquiries, fetch_qnas
from worker.naver_adapter import KST
from worker.naver_client import (
    API_BASE,
    ENV_FILE,
    TIMEOUT_SECONDS,
    fetch_token,
    load_credentials,
    read_env_file,
)
from worker.order_link import RetryingSession, parse_time

SWITCH_KEY = "CS_ANSWER_ENABLED"
QNA_ANSWER_URL = f"{API_BASE}/v1/contents/qnas/{{question_id}}"
INQUIRY_ANSWER_URL = f"{API_BASE}/v1/pay-merchant/inquiries/{{inquiry_no}}/answer"
# NOT verified against Naver's official docs: the Q&A field name is from a community report and the
# inquiry one is a best guess. Verify on the first live post; a 4xx reply shows Naver's own message.
QNA_BODY_FIELD = "commentContent"
INQUIRY_BODY_FIELD = "answerComment"
MAX_ANSWER_CHARS = 1000
CHANNEL_QNA, CHANNEL_INQUIRY = "상품문의", "고객문의"
ALREADY_ANSWERED = "이미 판매자센터에서 답변된 문의입니다."
CLIENT_ERROR_STATUS = 400


class CsAnswerError(RuntimeError):
    """The answer was not posted; the message is Korean and user-facing."""


class AlreadyAnswered(CsAnswerError):
    """The question already has an answer on Naver."""


@dataclass(frozen=True)
class AnswerResult:
    simulated: bool


def answer_enabled() -> bool:
    value = os.environ.get(SWITCH_KEY) or read_env_file(ENV_FILE).get(SWITCH_KEY, "")
    return value.strip() == "1"


def validate_answer(text: str) -> str:
    """The stripped answer; refuses empty or too long text (never truncates)."""
    cleaned = (text or "").strip()
    if not cleaned:
        raise CsAnswerError("답변이 비어 있습니다.")
    if len(cleaned) > MAX_ANSWER_CHARS:
        raise CsAnswerError(f"답변이 {len(cleaned)}자로 너무 깁니다. {MAX_ANSWER_CHARS}자 이하로 줄여 주세요.")
    return cleaned


def _write(method: Callable[..., requests.Response], url: str, headers: dict, body: dict) -> None:
    try:
        resp = method(url, headers=headers, json=body, timeout=TIMEOUT_SECONDS)
    except requests.RequestException as err:
        raise CsAnswerError(f"네이버에 연결하지 못했습니다 ({type(err).__name__}). 잠시 뒤 다시 시도해 주세요.") from None
    if resp.status_code < CLIENT_ERROR_STATUS:
        return
    try:
        data = resp.json()
        detail = f"{data.get('code', '')} {data.get('message', '')}".strip()
    except ValueError:
        detail = (getattr(resp, "text", "") or "")[:200]
    hint = " 요청 항목 이름이 맞지 않을 수 있으니 개발자에게 이 메시지를 알려 주세요." if resp.status_code < 500 else ""
    raise CsAnswerError(f"네이버가 답변 등록을 받아주지 않았습니다 (HTTP {resp.status_code}): {detail}.{hint}")


def post_qna_answer(token: str, question_id: str, text: str, session: requests.Session | None = None) -> None:
    http = RetryingSession(session)
    _write(http.put, QNA_ANSWER_URL.format(question_id=question_id),
           {"Authorization": f"Bearer {token}"}, {QNA_BODY_FIELD: text})


def post_inquiry_answer(token: str, inquiry_no: str, text: str, session: requests.Session | None = None) -> None:
    http = RetryingSession(session)
    _write(http.post, INQUIRY_ANSWER_URL.format(inquiry_no=inquiry_no),
           {"Authorization": f"Bearer {token}"}, {INQUIRY_BODY_FIELD: text})


def is_answered(channel: str, item_id: str, created: date, today: date, token: str,
                fetch_qna_list: Callable[..., list[dict]] = fetch_qnas,
                fetch_inquiry_list: Callable[..., list[dict]] = fetch_customer_inquiries) -> bool:
    """Fetch the question again; True when Naver says it is answered (unknown questions count as open)."""
    if channel == CHANNEL_QNA:
        found = [q for q in fetch_qna_list(token, created, today) if str(q.get("questionId")) == item_id]
    else:
        found = [i for i in fetch_inquiry_list(token, created, today) if str(i.get("inquiryNo")) == item_id]
    return any(bool(item.get("answered")) for item in found)


def submit_answer(channel: str, item_id: str, created_text: str, text: str, *,
                  enabled: bool | None = None, now: Callable[[], datetime] = lambda: datetime.now(KST),
                  token: str | None = None, session: requests.Session | None = None,
                  answered_check: Callable[..., bool] = is_answered,
                  qna_poster: Callable[..., None] = post_qna_answer,
                  inquiry_poster: Callable[..., None] = post_inquiry_answer) -> AnswerResult:
    """Check the text, refuse a question answered meanwhile, then post (or only pretend when the switch is off)."""
    cleaned = validate_answer(text)
    on = answer_enabled() if enabled is None else enabled
    access = token or fetch_token(load_credentials(), session)
    today = now().astimezone(KST).date()
    created = (parse_time(created_text) or now()).astimezone(KST).date()
    if answered_check(channel, item_id, created, today, access):
        raise AlreadyAnswered(ALREADY_ANSWERED)
    if not on:
        return AnswerResult(simulated=True)
    poster = qna_poster if channel == CHANNEL_QNA else inquiry_poster
    poster(access, item_id, cleaned, session)
    return AnswerResult(simulated=False)
