"""CS watcher cycles: baseline, alert once, state after each send, pairs, dry-run — no network (fakes)."""
import json
import logging
from datetime import datetime, timedelta
from types import SimpleNamespace

import pytest

from tests.test_cs_drafter import FakeClient, reply
from tests.test_order_link import PERSONAL, detail
from worker import cs_watch
from worker.cs_alert import NO_AI_NOTE
from worker.cs_drafter import DraftError
from worker.cs_watch import CsWatchError, Options, run_cycle
from worker.naver_adapter import KST

NOW = datetime(2026, 10, 10, 12, 0, tzinfo=KST)
CHATS = frozenset({111})


def iso(hours_ago):
    return (NOW - timedelta(hours=hours_ago)).isoformat()


def question(qid, hours_ago=1, answered=False, answer=None, text="언제 와요?"):
    return {"questionId": qid, "createDate": iso(hours_ago), "question": text, "answered": answered,
            "answer": answer, "productId": "555", "productName": "국산 풍천 장어", "maskedWriterId": "gotu****"}


def inquiry(no, hours_ago=1, answered=False, answer=None, po="2026100700001-1"):
    return {"inquiryNo": no, "inquiryRegistrationDateTime": iso(hours_ago), "inquiryContent": "택배 언제 오나요",
            "answered": answered, "answerContent": answer, "productOrderIdList": po,
            "productName": "국산 풍천 장어", "customerId": "gotu****", "customerName": "홍길동"}


class FakeTelegram:
    sent: list = []
    fail_on: set = set()
    crash_on: set = set()

    def __init__(self, token):
        pass

    def call(self, method, **params):
        text = params["text"]
        if any(marker in text for marker in FakeTelegram.crash_on):
            raise KeyError("crash")
        if any(marker in text for marker in FakeTelegram.fail_on) and params["chat_id"] == 111:
            raise RuntimeError("Telegram sendMessage: boom")
        FakeTelegram.sent.append((method, params))
        return {}


class World:
    """What the fake Naver API returns this cycle."""

    def __init__(self):
        self.qnas, self.inquiries, self.recent, self.direct = [], [], [], []


@pytest.fixture
def world(monkeypatch):
    w = World()
    FakeTelegram.sent, FakeTelegram.fail_on, FakeTelegram.crash_on = [], set(), set()
    monkeypatch.setattr(cs_watch, "fetch_token", lambda creds: "tok")
    monkeypatch.setattr(cs_watch, "load_credentials", lambda: None)
    monkeypatch.setattr(cs_watch, "fetch_qnas", lambda token, start, end: list(w.qnas))
    monkeypatch.setattr(cs_watch, "fetch_customer_inquiries", lambda token, start, end: list(w.inquiries))
    monkeypatch.setattr(cs_watch, "fetch_recent_details", lambda token, days, now=None: list(w.recent))
    monkeypatch.setattr(cs_watch, "fetch_details", lambda token, ids, session: list(w.direct))
    monkeypatch.setattr(cs_watch, "Telegram", FakeTelegram)
    monkeypatch.setattr(cs_watch, "BotSettings",
                        SimpleNamespace(load=lambda: SimpleNamespace(token="t", allowed_chats=CHATS)))
    return w


@pytest.fixture
def opts(tmp_path):
    return Options(tmp_path / "state")


def cycle(opts, client=None, now=NOW):
    return run_cycle(opts, client, "문서", "m", now=now)


def texts():
    return [p["text"] for _, p in FakeTelegram.sent]


def state_of(opts):
    return json.loads((opts.state_dir / "notified.json").read_text(encoding="utf-8"))


# --- first run / alert once ---------------------------------------------------------------------

def test_first_run_records_old_questions_silently_and_alerts_new_ones(world, opts, capsys):
    world.qnas = [question("old", hours_ago=10), question("new", hours_ago=1)]
    stats = cycle(opts)
    assert (stats.alerted, stats.skipped) == (1, 1)
    assert len(texts()) == 1 and FakeTelegram.sent[0][1]["disable_web_page_preview"] is True
    state = state_of(opts)
    assert state["qna:old"]["alerted"] is False and state["qna:new"]["alerted"] is True
    assert "알림 없이 기록만" in capsys.readouterr().out


def test_first_run_with_only_old_questions_still_creates_state_file(world, opts):
    world.qnas = [question("old", hours_ago=30)]
    assert cycle(opts).alerted == 0 and (opts.state_dir / "notified.json").exists()


def test_new_question_is_alerted_only_once(world, opts):
    world.qnas = [question("a")]
    cycle(opts)
    stats = cycle(opts)
    assert stats.alerted == 0 and len(texts()) == 1


def test_second_run_alerts_old_unseen_question(world, opts):
    cycle(opts)  # creates the (empty) state
    world.qnas = [question("late", hours_ago=20)]
    assert cycle(opts).alerted == 1


def test_already_answered_unknown_question_is_ignored(world, opts):
    world.qnas = [question("x", answered=True, answer="답")]
    cycle(opts)
    assert cycle(opts).alerted == 0 and texts() == [] and state_of(opts) == {}


def test_both_channels_alert_oldest_first_with_order_link(world, opts):
    world.qnas = [question("q", hours_ago=1)]
    world.inquiries = [inquiry("9", hours_ago=2)]
    world.recent = [detail()]
    cycle(opts)
    assert texts()[0].startswith("🔵 고객문의") and "주문: 정확" in texts()[0]
    assert texts()[1].startswith("🟡 상품문의")
    assert set(state_of(opts)) == {"qna:q", "inq:9"}


def test_inquiry_order_missing_from_recent_is_fetched_directly(world, opts):
    world.inquiries = [inquiry("9")]
    world.direct = [detail()]
    cycle(opts)
    assert "주문: 정확" in texts()[0]


def test_alert_has_no_personal_values(world, opts):
    world.inquiries = [inquiry("9")]
    world.recent = [detail()]
    cycle(opts)
    for secret in PERSONAL + ["2026100700001"]:
        assert secret not in texts()[0]


def test_no_ai_mode_note_and_null_draft_in_state(world, opts):
    world.qnas = [question("a")]
    cycle(opts, client=None)
    assert NO_AI_NOTE in texts()[0] and state_of(opts)["qna:a"]["draft"] is None


def test_ai_draft_is_in_alert_and_state(world, opts):
    world.qnas = [question("a")]
    cycle(opts, client=FakeClient(reply()))
    assert "안녕하세요 고객님^^" in texts()[0]
    assert state_of(opts)["qna:a"]["draft"]["route"] == "AI답변"


def test_draft_failure_still_alerts_with_reason(world, opts):
    world.qnas = [question("a")]
    cycle(opts, client=FakeClient(error=ConnectionError("down")))
    assert "(AI 초안 실패:" in texts()[0] and state_of(opts)["qna:a"]["alerted"] is True


# --- state after each send ----------------------------------------------------------------------

def test_telegram_failure_on_one_message_does_not_lose_others(world, opts):
    world.qnas = [question("a", text="첫째AAA"), question("b", hours_ago=0.5, text="둘째BBB")]
    FakeTelegram.fail_on = {"첫째AAA"}
    stats = cycle(opts)
    assert stats.alerted == 1 and stats.failed == 1
    assert set(state_of(opts)) == {"qna:b"}
    FakeTelegram.fail_on = set()
    assert cycle(opts).alerted == 1  # the failed one is retried next cycle
    assert set(state_of(opts)) == {"qna:a", "qna:b"}


def test_state_is_saved_after_each_send_even_if_a_later_one_crashes(world, opts):
    world.qnas = [question("a", text="첫째AAA"), question("b", hours_ago=0.5, text="둘째BBB")]
    FakeTelegram.crash_on = {"둘째BBB"}
    with pytest.raises(KeyError):
        cycle(opts)
    assert set(state_of(opts)) == {"qna:a"}


def test_one_chat_failing_still_counts_when_another_got_it(world, opts, monkeypatch):
    monkeypatch.setattr(cs_watch, "BotSettings", SimpleNamespace(
        load=lambda: SimpleNamespace(token="t", allowed_chats=frozenset({111, 222}))))
    world.qnas = [question("a", text="첫째AAA")]
    FakeTelegram.fail_on = {"첫째AAA"}
    assert cycle(opts).alerted == 1 and [p["chat_id"] for _, p in FakeTelegram.sent] == [222]


def test_no_allowed_chat_is_a_clear_error(world, opts, monkeypatch):
    monkeypatch.setattr(cs_watch, "BotSettings", SimpleNamespace(
        load=lambda: SimpleNamespace(token="t", allowed_chats=frozenset())))
    world.qnas = [question("a")]
    with pytest.raises(CsWatchError, match="TELEGRAM_ALLOWED_CHAT_IDS"):
        cycle(opts)


# --- pairs --------------------------------------------------------------------------------------

def pairs(opts):
    path = opts.state_dir / "pairs.jsonl"
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()] if path.exists() else []


def test_pair_recorded_once_when_owner_answers_later(world, opts):
    world.qnas = [question("a")]
    world.recent = [detail()]
    cycle(opts, client=FakeClient(reply()))
    world.qnas = [question("a", answered=True, answer="내일 도착합니다")]
    assert cycle(opts).paired == 1
    (pair,) = pairs(opts)
    assert pair["key"] == "qna:a" and pair["owner_answer"] == "내일 도착합니다"
    assert pair["draft"] == "안녕하세요 고객님^^" and pair["route"] == "AI답변" and pair["kind"] == "배송 일정"
    assert pair["question"] == "언제 와요?" and pair["order_kind"] == "추정" and pair["answered_seen_at"]
    assert state_of(opts)["qna:a"]["done"] is True
    assert cycle(opts).paired == 0 and len(pairs(opts)) == 1


def test_inquiry_pair_uses_answer_content(world, opts):
    world.inquiries = [inquiry("9")]
    world.recent = [detail()]
    cycle(opts, client=FakeClient(reply()))
    world.inquiries = [inquiry("9", answered=True, answer="오늘 도착")]
    cycle(opts)
    assert pairs(opts)[0]["owner_answer"] == "오늘 도착" and pairs(opts)[0]["order_kind"] == "정확"


def test_answered_question_without_draft_or_baseline_writes_no_pair_but_is_done(world, opts):
    world.qnas = [question("old", hours_ago=10), question("noai", hours_ago=1)]
    cycle(opts, client=None)
    world.qnas = [question("old", hours_ago=10, answered=True, answer="a"),
                  question("noai", hours_ago=1, answered=True, answer="b")]
    cycle(opts)
    assert pairs(opts) == [] and all(e["done"] for e in state_of(opts).values())


def test_answer_given_as_object_is_read(world, opts):
    world.qnas = [question("a")]
    cycle(opts, client=FakeClient(reply()))
    world.qnas = [question("a", answered=True, answer={"content": "객체 답변"})]
    cycle(opts)
    assert pairs(opts)[0]["owner_answer"] == "객체 답변"


# --- dry run ------------------------------------------------------------------------------------

def test_dry_run_prints_alerts_and_writes_nothing(world, tmp_path, capsys, monkeypatch):
    dry = Options(tmp_path / "dry", dry_run=True)
    monkeypatch.setattr(cs_watch, "BotSettings", SimpleNamespace(load=lambda: pytest.fail("no telegram")))
    world.qnas = [question("old", hours_ago=10), question("a")]
    stats = cycle(dry, client=FakeClient(reply()))
    out = capsys.readouterr().out
    assert stats.alerted == 1 and "🟡 상품문의" in out and "안녕하세요 고객님^^" in out
    assert not (tmp_path / "dry").exists() and FakeTelegram.sent == []


# --- loop and CLI -------------------------------------------------------------------------------

def test_watch_forever_survives_a_failed_cycle(monkeypatch, opts, caplog):
    outcomes = iter([RuntimeError("네이버 장애"), cs_watch.CycleStats(alerted=1)])

    def fake_cycle(*args):
        result = next(outcomes)
        if isinstance(result, Exception):
            raise result
        return result

    sleeps = []

    def fake_sleep(seconds):
        sleeps.append(seconds)
        if len(sleeps) == 2:
            raise KeyboardInterrupt

    monkeypatch.setattr(cs_watch, "run_cycle", fake_cycle)
    with caplog.at_level(logging.INFO, logger="cs_watch"), pytest.raises(KeyboardInterrupt):
        cs_watch.watch_forever(opts, 10, None, "", "m", sleep=fake_sleep)
    assert sleeps == [600, 600]
    assert "이번 주기 실패" in caplog.text and "알림 1건" in caplog.text


def test_default_state_dir_default_and_override(monkeypatch, tmp_path):
    monkeypatch.setattr(cs_watch, "ENV_FILE", tmp_path / "no.env")
    monkeypatch.delenv("CS_STATE_DIR", raising=False)
    default = cs_watch.default_state_dir()
    assert default.parts[-2:] == ("CS자료", "알림기록")
    monkeypatch.setenv("CS_STATE_DIR", str(tmp_path / "x"))
    assert cs_watch.default_state_dir() == tmp_path / "x"


@pytest.fixture
def cli(monkeypatch, tmp_path):
    monkeypatch.setattr(cs_watch, "default_state_dir", lambda: tmp_path / "s")
    monkeypatch.setattr(cs_watch, "model_name", lambda: "m")
    monkeypatch.setattr(cs_watch, "make_client", lambda: None)


def test_main_once_without_key_runs_in_no_ai_mode(cli, monkeypatch, capsys):
    seen = []
    monkeypatch.setattr(cs_watch, "run_cycle", lambda o, c, d, m: seen.append((o, c, d, m)) or cs_watch.CycleStats(2))
    assert cs_watch.main(["--once", "--days", "3", "--dry-run"]) == 0
    options, client, doc, model = seen[0]
    assert (options.days, options.dry_run, client, doc, model) == (3, True, None, "", "m")
    out = capsys.readouterr().out
    assert NO_AI_NOTE in out and "알림 2건" in out


def test_main_loads_doc_when_key_present_and_reports_missing_doc(cli, monkeypatch, capsys):
    client = object()
    monkeypatch.setattr(cs_watch, "make_client", lambda: client)
    monkeypatch.setattr(cs_watch, "load_store_doc", lambda: "문서내용")
    seen = []
    monkeypatch.setattr(cs_watch, "run_cycle", lambda o, c, d, m: seen.append((c, d)) or cs_watch.CycleStats())
    assert cs_watch.main(["--once"]) == 0 and seen == [(client, "문서내용")]

    def missing():
        raise DraftError("가게 정보 문서를 찾을 수 없습니다: x")

    monkeypatch.setattr(cs_watch, "load_store_doc", missing)
    assert cs_watch.main(["--once"]) == 1 and "찾을 수 없습니다" in capsys.readouterr().out


def test_main_once_failure_returns_1(cli, monkeypatch, capsys):
    def boom(*args):
        raise RuntimeError("네이버 오류")

    monkeypatch.setattr(cs_watch, "run_cycle", boom)
    assert cs_watch.main(["--once"]) == 1 and "확인 실패: 네이버 오류" in capsys.readouterr().out


def test_main_loop_mode_stops_on_ctrl_c(cli, monkeypatch):
    calls = []

    def fake_forever(opts, every, client, doc, model):
        calls.append(every)
        raise KeyboardInterrupt

    monkeypatch.setattr(cs_watch, "watch_forever", fake_forever)
    assert cs_watch.main(["--every", "5"]) == 0 and calls == [5]


# --- fetch window for late owner answers and doc reload ----------------------------------------

def test_fetch_start_reaches_back_to_oldest_waiting_alert():
    end = datetime(2026, 10, 10).date()
    state = {"qna:1": {"created": "2026-10-03T09:00:00+09:00", "alerted": True, "done": False},
             "qna:2": {"created": "2026-09-01T09:00:00+09:00", "alerted": True, "done": True},
             "qna:3": {"created": "2026-09-20T09:00:00+09:00", "alerted": False, "done": False}}
    assert cs_watch.fetch_start(state, end, 2) == datetime(2026, 10, 3).date()


def test_fetch_start_defaults_to_days_and_is_capped():
    end = datetime(2026, 10, 10).date()
    assert cs_watch.fetch_start({}, end, 2) == datetime(2026, 10, 9).date()
    old = {"qna:1": {"created": "2026-06-01T09:00:00+09:00", "alerted": True, "done": False}}
    assert cs_watch.fetch_start(old, end, 2) == end - timedelta(days=cs_watch.PAIR_WINDOW_DAYS - 1)


def test_watch_forever_rereads_the_store_doc_each_cycle(monkeypatch, opts):
    docs = iter(["v1", "v2"])
    seen = []
    monkeypatch.setattr(cs_watch, "load_store_doc", lambda: next(docs))
    monkeypatch.setattr(cs_watch, "run_cycle", lambda o, c, doc, m: seen.append(doc) or cs_watch.CycleStats())

    def stop_after_two(_seconds):
        if len(seen) == 2:
            raise KeyboardInterrupt
    with pytest.raises(KeyboardInterrupt):
        cs_watch.watch_forever(opts, 10, object(), "start", "m", sleep=stop_after_two)
    assert seen == ["v1", "v2"]


def test_inquiry_question_joins_title_and_body():
    assert cs_watch.inquiry_question({"title": "내일 몇 시에 오나요", "inquiryContent": "저녁에 먹을 수 있나요?"}) \
        == "내일 몇 시에 오나요\n저녁에 먹을 수 있나요?"
    assert cs_watch.inquiry_question({"title": " ", "inquiryContent": "본문만"}) == "본문만"
    assert cs_watch.inquiry_question({}) == ""
