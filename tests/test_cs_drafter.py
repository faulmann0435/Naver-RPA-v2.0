"""AI drafter: prompt building, masking, JSON validation, errors — no network (fake client)."""
import json
from datetime import datetime
from types import SimpleNamespace

import pytest

from tests.test_order_link import detail
from worker import cs_drafter
from worker.cs_drafter import (
    DraftError,
    DraftResult,
    InquiryView,
    build_messages,
    draft,
    load_store_doc,
    make_client,
    mask_question,
    parse_reply,
    stamp_with_weekday,
)
from worker.naver_adapter import KST
from worker.order_link import LinkResult, link_customer_inquiry

DOC = "# 가게 정보\n장어 예약발송: https://smartstore.naver.com/bestns/products/11368563340"
ASKED = datetime(2026, 10, 10, 9, 5, tzinfo=KST)  # a Saturday


def view(question="언제 와요?", link=None):
    link = link or link_customer_inquiry({"productOrderIdList": "2026100700001-1"}, [detail()])
    return InquiryView("상품문의", ASKED, "국산 풍천 장어", question, link)


class FakeClient:
    """Has the anthropic `messages.create` shape."""

    def __init__(self, text=None, error=None):
        self.calls, self._text, self._error = [], text, error
        self.messages = SimpleNamespace(create=self._create)

    def _create(self, **kwargs):
        self.calls.append(kwargs)
        if self._error:
            raise self._error
        return SimpleNamespace(content=[SimpleNamespace(type="text", text=self._text)])


def reply(**fields):
    base = {"kind": "배송 일정", "route": "AI답변", "reason": "문서에 있음", "draft": "안녕하세요 고객님^^"}
    return json.dumps({**base, **fields}, ensure_ascii=False)


# --- prompt ------------------------------------------------------------------------------------

def test_system_has_rules_then_cached_document():
    system = build_messages(DOC, view())["system"]
    assert "지어내지" in system[0]["text"] and "cache_control" not in system[0]
    assert DOC in system[1]["text"] and system[1]["cache_control"] == {"type": "ephemeral"}


def test_rules_cover_the_required_points():
    rules = build_messages(DOC, view())["system"][0]["text"]
    for phrase in ("요일", "정확", "§6", "§4", "주소 변경", "§5", "추가 구매", "주문번호", "링크"):
        assert phrase in rules


def test_user_message_has_channel_time_product_question_and_order_text():
    user = build_messages(DOC, view())["messages"][0]["content"]
    assert "채널: 상품문의" in user and "2026-10-10(토) 09:05" in user
    assert "상품: 국산 풍천 장어" in user and "문의 내용: 언제 와요?" in user
    assert "주문 연결: 정확" in user


def test_user_message_without_product_name():
    no_name = InquiryView("고객문의", ASKED, "", "문의", LinkResult("없음", (), "x"))
    assert "(알 수 없음)" in build_messages(DOC, no_name)["messages"][0]["content"]


def test_weekday_stamp():
    assert stamp_with_weekday(datetime(2026, 10, 12, 23, 59, tzinfo=KST)) == "2026-10-12(월) 23:59"


def test_masking_phones_and_long_numbers_but_not_product_links():
    text = "전화 010-1234-5678 또는 02.123.4567, 주문 2026100780314801 링크 products/11368563340 끝"
    out = mask_question(text)
    assert out.count("[전화]") == 2 and "[번호]" in out
    assert "5678" not in out and "2026100780314801" not in out
    assert "products/11368563340" in out
    assert mask_question("01012345678 입니다") == "[전화] 입니다"
    assert mask_question(None) == ""


def test_masked_question_reaches_prompt():
    user = build_messages(DOC, view("010 1234 5678로 연락 주세요"))["messages"][0]["content"]
    assert "5678" not in user and "[전화]" in user


# --- reply parsing -----------------------------------------------------------------------------

def test_parse_valid_reply():
    assert parse_reply(reply()) == DraftResult("배송 일정", "AI답변", "문서에 있음", "안녕하세요 고객님^^")


def test_parse_bad_route_falls_back_to_human_and_bad_kind_to_other():
    result = parse_reply(reply(route="자동발송", kind="이상한것"))
    assert result.route == "사람에게" and result.kind == "기타"


def test_parse_missing_draft_is_empty():
    assert parse_reply(json.dumps({"kind": "기타", "route": "사람에게"})).draft == ""


@pytest.mark.parametrize("text", ["not json", "[1, 2]", None])
def test_parse_garbage_raises_korean_error(text):
    with pytest.raises(DraftError, match="AI 응답"):
        parse_reply(text)


# --- draft() -----------------------------------------------------------------------------------

def test_draft_calls_api_with_schema_model_and_cache():
    client = FakeClient(reply())
    result = draft(client, DOC, view(), "claude-sonnet-5-5")
    call = client.calls[0]
    assert result.route == "AI답변" and call["model"] == "claude-sonnet-5-5"
    assert call["output_config"]["format"]["type"] == "json_schema"
    assert call["output_config"]["format"]["schema"]["properties"]["route"]["enum"] == ["AI답변", "사람에게"]
    assert call["system"][1]["cache_control"] == {"type": "ephemeral"}
    assert call["messages"][0]["role"] == "user"


def test_draft_api_error_becomes_module_error():
    client = FakeClient(error=ConnectionError("down"))
    with pytest.raises(DraftError, match="AI 초안 요청에 실패"):
        draft(client, DOC, view(), "m")


def test_draft_without_text_block_is_error():
    client = FakeClient("x")
    client.messages = SimpleNamespace(create=lambda **kw: SimpleNamespace(content=[]))
    with pytest.raises(DraftError):
        draft(client, DOC, view(), "m")


# --- document, model, client -------------------------------------------------------------------

def test_load_store_doc_reads_file_and_reports_missing_or_empty(tmp_path):
    path = tmp_path / "doc.md"
    path.write_text("  내용  ", encoding="utf-8")
    assert load_store_doc(path) == "내용"
    with pytest.raises(DraftError, match="찾을 수 없습니다"):
        load_store_doc(tmp_path / "none.md")
    path.write_text("  ", encoding="utf-8")
    with pytest.raises(DraftError, match="비어"):
        load_store_doc(path)


def test_doc_path_and_model_defaults_and_overrides(monkeypatch, tmp_path):
    monkeypatch.setattr(cs_drafter, "ENV_FILE", tmp_path / "no.env")
    monkeypatch.delenv("CS_STORE_DOC", raising=False)
    monkeypatch.delenv("CS_MODEL", raising=False)
    assert cs_drafter.doc_path().name == "가게정보_v2.md" and cs_drafter.doc_path().parent.name == "CS자료"
    assert cs_drafter.model_name() == "claude-sonnet-5-5"
    monkeypatch.setenv("CS_STORE_DOC", str(tmp_path / "x.md"))
    monkeypatch.setenv("CS_MODEL", "other-model")
    assert cs_drafter.doc_path() == tmp_path / "x.md" and cs_drafter.model_name() == "other-model"


def test_make_client_none_without_key(monkeypatch, tmp_path):
    monkeypatch.setattr(cs_drafter, "ENV_FILE", tmp_path / "no.env")
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    assert make_client() is None


def test_make_client_explains_a_missing_ai_package(monkeypatch):
    import builtins
    real_import = builtins.__import__

    def no_anthropic(name, *args, **kwargs):
        if name == "anthropic":
            raise ImportError(name)
        return real_import(name, *args, **kwargs)

    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    monkeypatch.setattr(builtins, "__import__", no_anthropic)
    with pytest.raises(DraftError, match="anthropic"):
        make_client()


def test_make_client_reads_env_file_and_builds_client(monkeypatch, tmp_path):
    env = tmp_path / ".env"
    env.write_text("ANTHROPIC_API_KEY=sk-test\n", encoding="utf-8")
    monkeypatch.setattr(cs_drafter, "ENV_FILE", env)
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    client = make_client()
    assert client is not None and hasattr(client.messages, "create")


def test_reply_cut_off_at_the_token_limit_is_a_draft_error():
    client = FakeClient(text='{"kind":"배송 일정","route":"AI답변","reason":"…","draft":"안녕')
    client.messages = SimpleNamespace(create=lambda **kw: SimpleNamespace(
        stop_reason="max_tokens", content=[SimpleNamespace(type="text", text='{"kind":"배송')]))
    with pytest.raises(DraftError, match="잘렸"):
        draft(client, DOC, view(), "m")
