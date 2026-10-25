"""CS watcher inside the bot: buttons, state entry, outcomes (answered elsewhere), the bot job."""
import json
import logging

import pytest

from store.memory_store import MemoryStore
from tests.test_cs_drafter import FakeClient, reply
from tests.test_cs_watch import (  # noqa: F401  (world and opts are pytest fixtures)
    CHATS,
    NOW,
    FakeTelegram,
    cycle,
    opts,
    pairs,
    question,
    state_of,
    texts,
    world,
)
from worker import cs_watch
from worker.cs_alert import NO_AI_NOTE
from worker.cs_drafter import DraftError
from worker.cs_state import short_id
from worker.cs_watch import Options, run_cycle


def outcomes_of(options):
    path = options.state_dir / "outcomes.jsonl"
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()] if path.exists() else []


def settle_in_state(options, **fields):
    state = state_of(options)
    cs_watch.save_state(options.state_dir, {**state, "qna:a": {**state["qna:a"], **fields}})


# --- alert buttons and the state entry -------------------------------------------------------------

def test_bot_mode_alert_has_buttons_and_state_keeps_what_the_flow_needs(world, tmp_path):
    bot_opts = Options(tmp_path / "state", buttons=True)
    world.qnas = [question("a")]
    cycle(bot_opts, client=FakeClient(reply()))
    ((_, params),) = FakeTelegram.sent
    data = [b["callback_data"] for row in params["reply_markup"]["inline_keyboard"] for b in row]
    sid = short_id("qna:a")
    assert data == [f"cs:ai:{sid}", f"cs:wr:{sid}", f"cs:sc:{sid}"]
    entry = state_of(bot_opts)["qna:a"]
    assert entry["sid"] == sid and entry["messages"] == {"111": 101}
    assert (entry["channel"], entry["id"], entry["product_name"], entry["question"]) == \
        ("상품문의", "a", "국산 풍천 장어", "언제 와요?")
    assert entry["created"] and entry["draft"]["draft"] == "안녕하세요 고객님^^"


def test_bot_mode_without_ai_has_only_write_and_seller_center_buttons(world, tmp_path):
    world.qnas = [question("a")]
    cycle(Options(tmp_path / "state", buttons=True), client=None)
    markup = FakeTelegram.sent[0][1]["reply_markup"]["inline_keyboard"]
    assert [row[0]["text"] for row in markup] == ["✍️ 직접 쓰기", "🙋 판매자센터에서 할게요"]
    assert NO_AI_NOTE in texts()[0]


def test_command_line_mode_sends_no_buttons(world, opts):
    world.qnas = [question("a")]
    cycle(opts, client=FakeClient(reply()))
    assert "reply_markup" not in FakeTelegram.sent[0][1]


def test_bot_passes_its_own_telegram_and_ai_note(world, tmp_path, monkeypatch):
    monkeypatch.setattr(cs_watch, "BotSettings", type("S", (), {"load": staticmethod(lambda: pytest.fail("own"))}))
    world.qnas = [question("a")]
    run_cycle(Options(tmp_path / "s"), None, "", "m", now=NOW, target=(FakeTelegram("t"), CHATS),
              ai_off_note="(AI 초안 꺼짐: 문서 없음)")
    assert "(AI 초안 꺼짐: 문서 없음)" in texts()[0]


# --- outcomes written by the watcher -----------------------------------------------------------------

def test_answered_in_seller_center_without_choice_is_answered_elsewhere(world, opts):
    world.qnas = [question("a", text="전화 010-1234-5678")]
    cycle(opts, client=FakeClient(reply()))
    world.qnas = [question("a", answered=True, answer="내일 도착합니다")]
    store = MemoryStore()
    run_cycle(opts, None, "", "m", now=NOW, store=store)
    (o,) = outcomes_of(opts)
    assert (o["outcome"], o["route"], o["owner_answer"]) == ("answered_elsewhere", "AI답변", "내일 도착합니다")
    assert o["ai_draft"] == "안녕하세요 고객님^^"
    remote = json.loads(store.read_text("cs/outcomes.json").content)
    assert "010-1234-5678" not in json.dumps(remote, ensure_ascii=False) and "[전화]" in remote[0]["question"]
    assert len(pairs(opts)) == 1  # the pair log is unchanged
    run_cycle(opts, None, "", "m", now=NOW, store=store)
    assert len(outcomes_of(opts)) == 1


def test_baseline_questions_get_no_outcome(world, opts):
    world.qnas = [question("old", hours_ago=10)]
    cycle(opts)
    world.qnas = [question("old", hours_ago=10, answered=True, answer="a")]
    cycle(opts)
    assert outcomes_of(opts) == []


def test_answer_posted_by_the_bot_is_not_paired_nor_counted_again(world, opts):
    world.qnas = [question("a")]
    cycle(opts, client=FakeClient(reply()))
    settle_in_state(opts, outcome="ai_as_is", outcome_simulated=False)
    world.qnas = [question("a", answered=True, answer="안녕하세요 고객님^^")]
    stats = cycle(opts)
    assert stats.paired == 0 and pairs(opts) == [] and outcomes_of(opts) == []
    assert state_of(opts)["qna:a"]["done"] is True


def test_simulated_post_then_real_owner_answer_still_makes_a_pair(world, opts):
    world.qnas = [question("a")]
    cycle(opts, client=FakeClient(reply()))
    settle_in_state(opts, outcome="ai_as_is", outcome_simulated=True)
    world.qnas = [question("a", answered=True, answer="진짜 답")]
    assert cycle(opts).paired == 1 and pairs(opts)[0]["owner_answer"] == "진짜 답" and outcomes_of(opts) == []


def test_seller_center_choice_then_owner_answer_makes_a_pair_only(world, opts):
    world.qnas = [question("a")]
    cycle(opts, client=FakeClient(reply()))
    settle_in_state(opts, outcome="seller_center", outcome_simulated=False)
    world.qnas = [question("a", answered=True, answer="센터 답")]
    assert cycle(opts).paired == 1 and outcomes_of(opts) == []


# --- make_bot_job ----------------------------------------------------------------------------------

@pytest.fixture
def job_env(monkeypatch):
    seen = []
    monkeypatch.setattr(cs_watch, "model_name", lambda: "m")
    monkeypatch.setattr(cs_watch, "make_client", lambda: "client")
    monkeypatch.setattr(cs_watch, "run_cycle", lambda options, client, doc, model, **kw:
                        seen.append((options, client, doc, kw)) or cs_watch.CycleStats())
    return seen


def test_bot_job_runs_a_button_cycle_with_the_store_doc(job_env, tmp_path):
    store = MemoryStore()
    store.write_text("cs/가게정보.md", "저장소 문서", None, "m")
    tg = object()
    cs_watch.make_bot_job(tg, CHATS, store, tmp_path)()
    ((options, client, doc, kw),) = job_env
    assert (options.buttons, options.state_dir, client, doc) == (True, tmp_path, "client", "저장소 문서")
    assert kw["target"] == (tg, CHATS) and kw["store"] is store


def test_bot_job_without_key_keeps_alerting_without_ai(job_env, tmp_path, monkeypatch):
    monkeypatch.setattr(cs_watch, "make_client", lambda: None)
    cs_watch.make_bot_job(object(), CHATS, None, tmp_path)()
    ((_, client, doc, kw),) = job_env
    assert client is None and doc == "" and kw["ai_off_note"] == NO_AI_NOTE


def test_bot_job_missing_doc_falls_back_to_no_ai_with_reason(job_env, tmp_path, monkeypatch, caplog):
    def missing(store=None):
        raise DraftError("가게 정보 문서를 찾을 수 없습니다: x")

    monkeypatch.setattr(cs_watch, "load_store_doc", missing)
    with caplog.at_level(logging.ERROR, logger="cs_watch"):
        cs_watch.make_bot_job(object(), CHATS, None, tmp_path)()
    ((_, client, _, kw),) = job_env
    assert client is None and "찾을 수 없습니다" in kw["ai_off_note"] and "문서" in caplog.text


def test_bot_job_rereads_the_doc_every_cycle(job_env, tmp_path, monkeypatch):
    docs = iter(["v1", "v2"])
    monkeypatch.setattr(cs_watch, "load_store_doc", lambda store=None: next(docs))
    job = cs_watch.make_bot_job(object(), CHATS, None, tmp_path)
    job()
    job()
    assert [d for _, _, d, _ in job_env] == ["v1", "v2"]


def test_bot_job_flushes_the_outcome_outbox(job_env, tmp_path, monkeypatch):
    flushed = []
    monkeypatch.setattr(cs_watch, "flush_outbox", lambda directory, store: flushed.append((directory, store)))
    store = MemoryStore()
    store.write_text("cs/가게정보.md", "d", None, "m")
    cs_watch.make_bot_job(object(), CHATS, store, tmp_path)()
    assert flushed == [(tmp_path, store)]


def test_bot_job_survives_a_broken_ai_setup(job_env, tmp_path, monkeypatch):
    def boom():
        raise ImportError("anthropic")

    monkeypatch.setattr(cs_watch, "make_client", boom)
    cs_watch.make_bot_job(object(), CHATS, None, tmp_path)()
    assert job_env[0][1] is None
    assert job_env[0][3]["ai_off_note"] != NO_AI_NOTE and "ImportError" in job_env[0][3]["ai_off_note"]


def test_bot_job_names_the_ai_setup_problem(job_env, tmp_path, monkeypatch):
    def missing_package():
        raise DraftError("AI 부품(anthropic)이 설치되지 않았습니다")

    monkeypatch.setattr(cs_watch, "make_client", missing_package)
    cs_watch.make_bot_job(object(), CHATS, None, tmp_path)()
    assert "anthropic" in job_env[0][3]["ai_off_note"]


def test_bot_job_picks_up_a_key_added_later(job_env, tmp_path, monkeypatch):
    clients = iter([None, "client"])
    monkeypatch.setattr(cs_watch, "make_client", lambda: next(clients))
    store = MemoryStore()
    store.write_text("cs/가게정보.md", "d", None, "m")
    job = cs_watch.make_bot_job(object(), CHATS, store, tmp_path)
    job()
    job()
    assert [seen[1] for seen in job_env] == [None, "client"]
