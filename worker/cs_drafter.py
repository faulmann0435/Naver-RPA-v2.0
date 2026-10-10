"""AI draft answer for one customer question, based ONLY on the store document (read-only).

The draft is never posted anywhere: it only goes to the owner's phone (see worker.cs_watch).
Claude is asked through the Messages API; the long store document sits in the system prompt with a
cache breakpoint (so repeated questions reuse it), and a JSON schema (`output_config.format`)
makes the reply machine-readable.

.env keys: ANTHROPIC_API_KEY, CS_MODEL (default below), CS_STORE_DOC (path of the store document).
"""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from tests.regression_harness import REPO_ROOT
from worker.naver_adapter import KST
from worker.naver_client import ENV_FILE, read_env_file
from worker.order_link import LinkResult, to_prompt_text

DEFAULT_MODEL = "claude-sonnet-5-5"
MODEL_KEY, DOC_KEY, API_KEY = "CS_MODEL", "CS_STORE_DOC", "ANTHROPIC_API_KEY"
DEFAULT_DOC = REPO_ROOT.parent / "CS자료" / "가게정보_v2.md"
MAX_TOKENS = 4096  # the model thinks before answering; 1024 cut replies off mid-JSON
KINDS = ("배송 일정", "상품 정보", "보관·조리", "취소·변경", "불만·클레임", "기타")
ROUTE_AI, ROUTE_HUMAN = "AI답변", "사람에게"
ROUTES = (ROUTE_AI, ROUTE_HUMAN)
WEEKDAYS = "월화수목금토일"

PHONE = re.compile(r"(?<![\d/])0\d{1,2}[-. ]?\d{3,4}[-. ]?\d{4}(?!\d)")
LONG_DIGITS = re.compile(r"(?<!\d)(?<!products/)\d{10,}")  # product links keep their number

RULES = """당신은 수산물 스마트스토어 '최고다농수산'의 고객 문의에 대한 답변 '초안'을 쓰는 도우미입니다.
초안은 사장님만 보고, 사장님이 고쳐서 직접 답합니다. 아래 규칙을 반드시 지킵니다.

1. 근거는 아래 [가게 정보 문서]와 주문 정보뿐입니다. 문서에 없는 내용은 지어내지 마세요.
2. 날짜 계산은 문의 시각(요일 포함)을 기준으로 합니다.
3. 주문 연결이 "정확"이면 결제 시각·주문 상태를 근거로 답해도 됩니다.
   "추정", "후보 여러 건", "없음"이면 출고일·도착일을 단정하지 마세요.
4. 문서 §6 '오늘의 상황' 칸에 없는 그날의 상황(조업·재고·입고·지연 등)은 확답하지 말고 route를 "사람에게"로 하세요.
5. 문서 §4에 해당하면 route를 "사람에게"로 하세요. 주소 변경·취소 요청, 주문 후 포장·구성·소스 변경 요청은
   언제나 "사람에게"이며, 초안에서 가능하다고 약속하지도 불가능하다고 단정하지도 마세요.
6. 말투는 문서 §5를 따르세요. 묻지 않은 것은 덧붙이지 말고, 추가 구매를 권하지 마세요.
7. 고객 이름·전화번호·주소·주문번호를 답변에 쓰지 마세요.
8. 문서의 링크는 한 글자도 바꾸지 말고 그대로 쓰세요.
9. route가 "사람에게"여도 draft에는 사장님이 참고할 수 있는 초안을 쓰되, 확정할 수 없는 부분은 "확인 후 안내드리겠습니다"로 두세요.
   초안을 쓸 수 없으면 draft는 빈 문자열로 두세요.

출력은 JSON 하나입니다.
- kind: 문의 종류 (배송 일정 / 상품 정보 / 보관·조리 / 취소·변경 / 불만·클레임 / 기타)
- route: "AI답변"(문서만으로 안전하게 답할 수 있음) 또는 "사람에게"
- reason: 그렇게 판단한 이유 한 줄
- draft: 고객에게 보낼 답변 초안"""

SCHEMA = {
    "type": "object",
    "properties": {
        "kind": {"type": "string", "enum": list(KINDS)},
        "route": {"type": "string", "enum": list(ROUTES)},
        "reason": {"type": "string"},
        "draft": {"type": "string"},
    },
    "required": ["kind", "route", "reason", "draft"],
    "additionalProperties": False,
}


class DraftError(RuntimeError):
    """The store document is missing or the AI call failed; the message is Korean and user-facing."""


@dataclass(frozen=True)
class DraftResult:
    kind: str
    route: str
    reason: str
    draft: str


@dataclass(frozen=True)
class InquiryView:
    """What the drafter and the alert may know about one question (no personal values)."""
    channel: str  # "상품문의" | "고객문의"
    created: datetime
    product_name: str
    question: str
    link: LinkResult


def _setting(key: str) -> str:
    return os.environ.get(key) or read_env_file(ENV_FILE).get(key, "")


def model_name() -> str:
    return _setting(MODEL_KEY) or DEFAULT_MODEL


def doc_path() -> Path:
    override = _setting(DOC_KEY)
    return Path(override) if override else DEFAULT_DOC


def load_store_doc(path: Path | None = None) -> str:
    target = path or doc_path()
    if not target.is_file():
        raise DraftError(f"가게 정보 문서를 찾을 수 없습니다: {target} (.env의 {DOC_KEY} 확인)")
    text = target.read_text(encoding="utf-8").strip()
    if not text:
        raise DraftError(f"가게 정보 문서가 비어 있습니다: {target}")
    return text


def mask_question(text: str) -> str:
    """Hide phone numbers and long digit runs (order numbers); product link numbers stay."""
    return LONG_DIGITS.sub("[번호]", PHONE.sub("[전화]", text or "")).strip()


def stamp_with_weekday(moment: datetime) -> str:
    local = moment.astimezone(KST)
    return f"{local:%Y-%m-%d}({WEEKDAYS[local.weekday()]}) {local:%H:%M}"


def build_messages(doc: str, view: InquiryView) -> dict[str, Any]:
    """`system` (rules + cached document) and `messages` for the Messages API."""
    system = [
        {"type": "text", "text": RULES},
        {"type": "text", "text": f"[가게 정보 문서]\n{doc}", "cache_control": {"type": "ephemeral"}},
    ]
    user = "\n".join([
        f"채널: {view.channel}",
        f"문의 시각: {stamp_with_weekday(view.created)}",
        f"상품: {view.product_name or '(알 수 없음)'}",
        f"문의 내용: {mask_question(view.question)}",
        "",
        to_prompt_text(view.link),
    ])
    return {"system": system, "messages": [{"role": "user", "content": user}]}


def parse_reply(text: str) -> DraftResult:
    """Validate the model's JSON: unknown route -> 사람에게, unknown kind -> 기타, no draft -> ''."""
    try:
        data = json.loads(text)
    except (TypeError, ValueError) as err:
        raise DraftError("AI 응답을 JSON으로 읽을 수 없습니다.") from err
    if not isinstance(data, dict):
        raise DraftError("AI 응답 형식이 올바르지 않습니다.")
    kind = data.get("kind") if data.get("kind") in KINDS else "기타"
    route = data.get("route") if data.get("route") in ROUTES else ROUTE_HUMAN
    return DraftResult(kind, route, str(data.get("reason") or ""), str(data.get("draft") or ""))


def draft(client: Any, doc: str, view: InquiryView, model: str) -> DraftResult:
    """Ask Claude for a draft. `client` only needs the anthropic `messages.create` shape."""
    try:
        reply = client.messages.create(model=model, max_tokens=MAX_TOKENS,
                                       output_config={"format": {"type": "json_schema", "schema": SCHEMA}},
                                       **build_messages(doc, view))
    except Exception as err:  # SDK errors (auth, rate limit, network) all mean: no draft this time
        raise DraftError(f"AI 초안 요청에 실패했습니다: {type(err).__name__}") from err
    if getattr(reply, "stop_reason", None) == "max_tokens":
        raise DraftError("AI 응답이 길이 한도에서 잘렸습니다.")
    text = next((b.text for b in reply.content if getattr(b, "type", "") == "text"), "")
    return parse_reply(text)


def make_client():
    """Anthropic client from ANTHROPIC_API_KEY (env or .env); None when there is no key."""
    key = _setting(API_KEY)
    if not key:
        return None
    import anthropic
    return anthropic.Anthropic(api_key=key)
