"""CS buttons and typed answers through the bot: confirm/post, write/preview, reply, cancel, guards."""
import json
from datetime import datetime, timedelta

import pytest

from store.memory_store import MemoryStore
from worker import cs_flow
from worker.cs_alert import alert_keyboard, preview_keyboard
from worker.cs_answer import AlreadyAnswered, AnswerResult, CsAnswerError
from worker.cs_drafter import DraftResult
from worker.cs_flow import CsFlow, fingerprint
from worker.cs_state import (
    OUTCOMES_FILE,
    PENDING_FILE,
    find_by_sid,
    load_state,
    read_jsonl,
    save_state,
    short_id,
)
from worker.naver_adapter import KST
from worker.telegram_bot import OrderBot

OWNER = 111
NOW = datetime(2026, 10, 20, 12, 0, tzinfo=KST)
KEY = "qna:5"
SID = short_id(KEY)
DRAFT = "안녕하세요 고객님^^ 내일 도착합니다."


def entry(draft=DRAFT, route="AI답변", **kw):
    d = {"kind": "배송 일정", "route": route, "reason": "r", "draft": draft} if draft is not None else None
    return {"sid": SID, "channel": "상품문의", "id": "5", "created": "2026-10-20T09:00:00+09:00",
            "product_name": "가짜 장어", "question": "언제 와요?", "draft": d, "alerted": True, "done": False,
            "messages": {str(OWNER): 77}, **kw}


class FakeTelegram:
    def __init__(self):
        self.sent = []

    def call(self, method, files=None, **params):
        self.sent.append((method, params))
        return {}

    def send(self, chat_id, text, **extra):
        self.sent.append(("sendMessage", {"chat_id": chat_id, "text": text, **extra}))

    def texts(self):
        return [p["text"] for m, p in self.sent if m == "sendMessage"]

    def cleared(self):
        return [p["message_id"] for m, p in self.sent if m == "editMessageReplyMarkup"]


class Poster:
    def __init__(self, result=AnswerResult(simulated=False), error=None):
        self.calls, self.result, self.error = [], result, error

    def __call__(self, entry, text):
        self.calls.append((entry["id"], text))
        if self.error:
            raise self.error
        return self.result


class Clock:
    def __init__(self):
        self.now = NOW

    def __call__(self):
        return self.now


@pytest.fixture
def env(tmp_path):
    save_state(tmp_path, {KEY: entry()})
    tg, poster, clock, store = FakeTelegram(), Poster(), Clock(), MemoryStore()
    flow = CsFlow(tg, tmp_path, store, answer_fn=poster, now=clock)
    bot = OrderBot(tg, frozenset({OWNER}), pending_file=tmp_path / "p.json", cs=flow)
    return type("Env", (), dict(tg=tg, poster=poster, clock=clock, store=store, flow=flow, bot=bot, dir=tmp_path))


def press(env, data, chat=OWNER):
    env.bot.handle({"callback_query": {"id": "q", "data": data,
                                       "message": {"chat": {"id": chat}, "message_id": 5}}})


def say(env, text, chat=OWNER, reply_to=None):
    message = {"chat": {"id": chat}, "text": text}
    if reply_to is not None:
        message["reply_to_message"] = {"message_id": reply_to}
    env.bot.handle({"message": message})


def last_buttons(env):
    markup = env.tg.sent[-1][1]["reply_markup"]["inline_keyboard"]
    return [b["callback_data"] for row in markup for b in row]


def outcomes(env):
    return read_jsonl(env.dir, OUTCOMES_FILE)


def post_button(env):
    return next(d for d in last_buttons(env) if d.startswith(f"cs:go:{SID}"))


# --- keyboards ----------------------------------------------------------------------------------

def test_alert_keyboard_per_route_and_draft():
    ai = alert_keyboard(SID, DraftResult("배송 일정", "AI답변", "r", "초안"))
    human = alert_keyboard(SID, DraftResult("불만·클레임", "사람에게", "r", "초안"))
    no_text = alert_keyboard(SID, DraftResult("기타", "사람에게", "r", "  "))
    none = alert_keyboard(SID, None)
    labels = lambda kb: [b["text"] for row in kb["inline_keyboard"] for b in row]
    assert labels(ai) == labels(human) == ["✅ 이대로 올리기", "✍️ 직접 쓰기", "🙋 판매자센터에서 할게요"]
    assert labels(no_text) == labels(none) == ["✍️ 직접 쓰기", "🙋 판매자센터에서 할게요"]
    for kb in (ai, human, none):
        for row in kb["inline_keyboard"]:
            for b in row:
                assert len(b["callback_data"].encode()) <= 64 and b["callback_data"].startswith("cs:")


def test_preview_keyboard_has_rewrite_only_for_owner_text():
    labels = lambda kb: [b["text"] for row in kb["inline_keyboard"] for b in row]
    assert labels(preview_keyboard(SID, "abc123", rewrite=False)) == ["올리기", "취소"]
    assert labels(preview_keyboard(SID, "abc123", rewrite=True)) == ["올리기", "✍️ 다시 쓰기", "취소"]
    data = preview_keyboard(SID, fingerprint("x"), rewrite=True)["inline_keyboard"][0][0]["callback_data"]
    assert len(data.encode()) <= 64


# --- ✅ post the AI draft -----------------------------------------------------------------------

def test_ai_draft_confirm_then_post(env):
    press(env, f"cs:ai:{SID}")
    assert env.tg.texts()[-1].startswith("이 답변을 올릴까요?") and DRAFT in env.tg.texts()[-1]
    assert env.poster.calls == []
    press(env, post_button(env))
    assert env.poster.calls == [("5", DRAFT)]
    assert "네이버에 답변을 올렸습니다" in env.tg.texts()[-1]
    assert env.tg.cleared() == [77]
    (o,) = outcomes(env)
    assert (o["outcome"], o["simulated"], o["route"], o["owner_answer"]) == ("ai_as_is", False, "AI답변", "")
    assert find_by_sid(load_state(env.dir)[0], SID)[1]["outcome"] == "ai_as_is"
    assert len(json.loads(env.store.read_text("cs/outcomes.json").content)) == 1


def test_simulated_post_says_test_mode_and_is_recorded_as_simulated(env):
    env.poster.result = AnswerResult(simulated=True)
    press(env, f"cs:ai:{SID}")
    press(env, post_button(env))
    assert "시험 모드: 네이버에 올리지 않았습니다" in env.tg.texts()[-1]
    assert outcomes(env)[0]["simulated"] is True


def test_double_tap_on_post_posts_once(env):
    press(env, f"cs:ai:{SID}")
    data = post_button(env)
    press(env, data)
    press(env, data)
    assert len(env.poster.calls) == 1 and len(outcomes(env)) == 1
    assert env.tg.texts()[-1] in (cs_flow.SETTLED, cs_flow.EXPIRED)


def test_tapping_ai_button_again_after_settling_is_refused(env):
    press(env, f"cs:ai:{SID}")
    press(env, post_button(env))
    press(env, f"cs:ai:{SID}")
    assert env.tg.texts()[-1] == cs_flow.SETTLED and len(env.poster.calls) == 1


def test_no_draft_text_means_nothing_to_post(env):
    save_state(env.dir, {KEY: entry(draft=None)})
    press(env, f"cs:ai:{SID}")
    assert "초안이 없습니다" in env.tg.texts()[-1] and env.poster.calls == []


def test_cancel_removes_the_preview(env):
    press(env, f"cs:ai:{SID}")
    data = post_button(env)
    press(env, f"cs:no:{SID}")
    press(env, data)
    assert env.poster.calls == [] and env.tg.texts()[-1] == cs_flow.EXPIRED


def test_already_answered_is_reported_and_settled_as_elsewhere(env):
    env.poster.error = AlreadyAnswered("이미 판매자센터에서 답변된 문의입니다.")
    press(env, f"cs:ai:{SID}")
    press(env, post_button(env))
    assert "이미 판매자센터에서 답변된 문의입니다" in env.tg.texts()[-1]
    assert outcomes(env)[0]["outcome"] == "answered_elsewhere" and env.tg.cleared() == [77]


@pytest.mark.parametrize("error", [CsAnswerError("네이버가 거절 (HTTP 400): X"), RuntimeError("boom")])
def test_failed_post_reports_and_offers_a_retry(env, error):
    env.poster.error = error
    press(env, f"cs:ai:{SID}")
    press(env, post_button(env))
    texts = env.tg.texts()
    assert "올리지 못했습니다" in texts[-2] and texts[-1].startswith("이 답변을 올릴까요?")
    assert outcomes(env) == [] and post_button(env)
    env.poster.error = None
    press(env, post_button(env))
    assert len(outcomes(env)) == 1


def test_stale_fingerprint_cannot_post_a_newer_text(env):
    press(env, f"cs:ai:{SID}")
    press(env, f"cs:go:{SID}:ffffff")
    assert env.poster.calls == [] and env.tg.texts()[-1] == cs_flow.EXPIRED


# --- ✍️ write your own --------------------------------------------------------------------------

def test_write_then_text_then_preview_then_post(env):
    press(env, f"cs:wr:{SID}")
    assert env.tg.texts()[-1] == "답변을 입력해 주세요 (취소하려면 /취소)"
    say(env, "직접 쓴 답변입니다")
    assert env.tg.texts()[-1].startswith("이렇게 올릴까요?") and "직접 쓴 답변입니다" in env.tg.texts()[-1]
    assert any(d.startswith(f"cs:rw:{SID}") for d in last_buttons(env))
    press(env, post_button(env))
    assert env.poster.calls == [("5", "직접 쓴 답변입니다")]
    (o,) = outcomes(env)
    assert (o["outcome"], o["owner_answer"], o["ai_draft"]) == ("owner_written", "직접 쓴 답변입니다", DRAFT)
    say(env, "그냥 말")  # waiting is over: normal help reply, no new preview
    assert "발주 처리" in env.tg.texts()[-1]


def test_owner_typing_the_draft_text_is_still_owner_written(env):
    press(env, f"cs:wr:{SID}")
    say(env, DRAFT)
    press(env, post_button(env))
    assert outcomes(env)[0]["outcome"] == "owner_written"


def test_reply_to_the_alert_works_without_pressing_write(env):
    say(env, "답장으로 쓴 답변", reply_to=77)
    assert env.tg.texts()[-1].startswith("이렇게 올릴까요?")
    press(env, post_button(env))
    assert env.poster.calls == [("5", "답장으로 쓴 답변")]


def test_reply_to_some_other_message_is_a_normal_message(env):
    say(env, "아무 말", reply_to=999)
    assert "발주 처리" in env.tg.texts()[-1] and env.poster.calls == []


def test_rewrite_button_waits_again(env):
    press(env, f"cs:wr:{SID}")
    say(env, "첫 번째")
    old = post_button(env)
    press(env, f"cs:rw:{SID}")
    assert env.tg.texts()[-1].startswith("답변을 입력해 주세요")
    say(env, "두 번째")
    new = post_button(env)
    press(env, old)  # the first preview is gone
    assert env.poster.calls == []
    press(env, new)
    assert env.poster.calls == [("5", "두 번째")]


def test_cancel_command_stops_waiting(env):
    press(env, f"cs:wr:{SID}")
    say(env, "/취소")
    assert env.tg.texts()[-1] == "취소했습니다."
    say(env, "다음 말")
    assert "발주 처리" in env.tg.texts()[-1] and env.poster.calls == []


def test_cancel_command_without_waiting_is_a_normal_message(env):
    say(env, "/취소")
    assert "발주 처리" in env.tg.texts()[-1]


def test_too_long_text_keeps_waiting(env):
    press(env, f"cs:wr:{SID}")
    say(env, "가" * 1001)
    assert "너무 깁니다" in env.tg.texts()[-1]
    say(env, "짧게 고침")
    assert env.tg.texts()[-1].startswith("이렇게 올릴까요?")


def test_commands_and_run_buttons_are_not_swallowed_while_waiting(env):
    ran = []
    env.bot.runner = lambda: ran.append(1) or (_ for _ in ()).throw(RuntimeError("stop"))
    press(env, f"cs:wr:{SID}")
    say(env, "/start")
    assert "발주 처리" in env.tg.texts()[-1] and env.poster.calls == []
    say(env, "📦 발주 처리")
    assert ran == [1]  # went to the order flow, not taken as an answer
    say(env, "📊 CS 성적")
    assert env.tg.texts()[-1].startswith("📊 CS 성적")
    say(env, "이건 답변")  # still waiting afterwards
    assert env.tg.texts()[-1].startswith("이렇게 올릴까요?")


# --- 🙋 seller center ---------------------------------------------------------------------------

def test_seller_center_choice_records_and_clears_buttons(env):
    press(env, f"cs:sc:{SID}")
    assert "판매자센터" in env.tg.texts()[-1] and env.tg.cleared() == [77]
    (o,) = outcomes(env)
    assert o["outcome"] == "seller_center" and env.poster.calls == []
    press(env, f"cs:ai:{SID}")
    assert env.tg.texts()[-1] == cs_flow.SETTLED


# --- guards -------------------------------------------------------------------------------------

def test_disallowed_chat_can_do_nothing(env):
    press(env, f"cs:ai:{SID}", chat=999)
    say(env, "남의 답변", chat=999, reply_to=77)
    assert env.poster.calls == [] and outcomes(env) == []
    assert not (env.dir / PENDING_FILE).exists()
    assert all("남의 답변" not in t for t in env.tg.texts())


def test_unknown_sid_is_expired(env):
    press(env, "cs:ai:deadbeef")
    assert env.tg.texts()[-1] == cs_flow.EXPIRED


def test_cs_off_ignores_cs_buttons_and_commands(env):
    env.bot.cs = None
    press(env, f"cs:ai:{SID}")
    assert env.tg.texts() == [] and [m for m, _ in env.tg.sent] == ["answerCallbackQuery"]
    say(env, "/성적")
    assert "발주 처리" in env.tg.texts()[-1]


# --- persistence --------------------------------------------------------------------------------

def test_waiting_survives_a_restart(env):
    press(env, f"cs:wr:{SID}")
    env.bot.cs = CsFlow(env.tg, env.dir, env.store, answer_fn=env.poster, now=env.clock)
    say(env, "재시작 뒤 답변")
    assert env.tg.texts()[-1].startswith("이렇게 올릴까요?")


def test_waiting_expires_after_24_hours(env):
    press(env, f"cs:wr:{SID}")
    env.clock.now = NOW + timedelta(hours=24, minutes=1)
    say(env, "너무 늦은 답변")
    assert "발주 처리" in env.tg.texts()[-1] and env.poster.calls == []


def test_preview_expires_after_24_hours(env):
    press(env, f"cs:ai:{SID}")
    data = post_button(env)
    env.clock.now = NOW + timedelta(hours=25)
    press(env, data)
    assert env.poster.calls == [] and env.tg.texts()[-1] == cs_flow.EXPIRED


def test_broken_pending_file_counts_as_empty(env):
    (env.dir / PENDING_FILE).write_text("{not json", encoding="utf-8")
    say(env, "아무 말")
    assert "발주 처리" in env.tg.texts()[-1]


def test_store_down_does_not_break_the_flow(env):
    class Broken(MemoryStore):
        def read_text(self, path):
            raise RuntimeError("down")

    env.flow.store = Broken()
    press(env, f"cs:ai:{SID}")
    press(env, post_button(env))
    assert "네이버에 답변을 올렸습니다" in env.tg.texts()[-1] and len(outcomes(env)) == 1


# --- stats command ------------------------------------------------------------------------------

def test_stats_command_and_button_show_numbers(env):
    press(env, f"cs:sc:{SID}")
    for text in ("/성적", "📊 CS 성적"):
        say(env, text)
        assert env.tg.texts()[-1].startswith("📊 CS 성적") and "🙋 판매자센터 1건 (100%)" in env.tg.texts()[-1]
