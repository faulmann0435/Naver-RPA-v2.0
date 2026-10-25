"""CS inside the bot process: the periodic job in poll_forever, the switch, the stats button, store document order."""
import pytest

from store.memory_store import MemoryStore
from worker import cs_drafter, cs_state, telegram_bot
from worker.cs_drafter import DraftError, load_store_doc
from worker.telegram_bot import (
    CS_INTERVAL,
    STATS_BUTTON,
    OrderBot,
    poll_forever,
    run_keyboard,
)


class _Stop(Exception):
    pass


def _poll_with_jobs(cs_job, relay=None, polls=9, step=100.0):
    clock = {"t": 1000.0}
    calls: list[float] = []

    class Tg:
        n = 0

        def call(self, method, **params):
            Tg.n += 1
            clock["t"] += step
            if Tg.n > polls:
                raise _Stop()
            return []

    job = cs_job(clock, calls)
    with pytest.raises(_Stop):
        poll_forever(OrderBot(Tg(), frozenset()), sleep=lambda s: None, relay=relay,
                     clock=lambda: clock["t"], cs_job=job)
    return calls


def test_cs_interval_is_ten_minutes():
    assert CS_INTERVAL == 600


def test_cs_job_runs_at_start_then_every_interval():
    calls = _poll_with_jobs(lambda clock, calls: lambda: calls.append(clock["t"]), polls=14)
    assert calls[0] == 1000.0
    assert all(b - a >= CS_INTERVAL for a, b in zip(calls, calls[1:])) and len(calls) == 3


def test_cs_job_errors_are_logged_and_never_stop_the_bot(caplog):
    def make(clock, calls):
        def job():
            calls.append(clock["t"])
            raise RuntimeError("naver down")
        return job

    with caplog.at_level("ERROR", logger="telegram_bot"):
        calls = _poll_with_jobs(make, polls=14)
    assert len(calls) == 3 and any("고객 문의 확인 실패" in r.message for r in caplog.records)


def test_cs_job_and_relay_keep_their_own_timing():
    relay_calls: list[float] = []
    cs_calls: list[float] = []

    def make(clock, calls):
        return lambda: cs_calls.append(clock["t"])

    clock_relay = lambda: relay_calls.append(1)
    _poll_with_jobs(make, relay=clock_relay, polls=14)
    assert len(relay_calls) > len(cs_calls) >= 2


def test_no_cs_job_means_poll_works_as_before():
    clock = {"t": 0.0}

    class Tg:
        n = 0

        def call(self, method, **params):
            Tg.n += 1
            if Tg.n > 2:
                raise _Stop()
            return []

    with pytest.raises(_Stop):
        poll_forever(OrderBot(Tg(), frozenset()), sleep=lambda s: None, clock=lambda: clock["t"])


# --- switch and keyboard ------------------------------------------------------------------------

def test_cs_enabled_reads_env_then_dotenv(monkeypatch, tmp_path):
    env = tmp_path / ".env"
    monkeypatch.setattr(cs_state, "ENV_FILE", env)
    monkeypatch.delenv("CS_WATCH_ENABLED", raising=False)
    assert cs_state.cs_enabled() is False
    env.write_text("CS_WATCH_ENABLED=1\n", encoding="utf-8")
    assert cs_state.cs_enabled() is True
    monkeypatch.setenv("CS_WATCH_ENABLED", "0")
    assert cs_state.cs_enabled() is False


def test_stats_button_is_on_the_keyboard_only_when_cs_is_on(monkeypatch):
    monkeypatch.setattr(telegram_bot, "cs_enabled", lambda: False)
    assert run_keyboard()["keyboard"] == [[{"text": "📦 발주 처리"}, {"text": "📋 발송 안 된 주문 전부"}]]
    monkeypatch.setattr(telegram_bot, "cs_enabled", lambda: True)
    assert run_keyboard()["keyboard"][1] == [{"text": STATS_BUTTON}]


def test_build_cs_is_off_without_the_switch(monkeypatch, caplog):
    monkeypatch.setattr(telegram_bot, "cs_enabled", lambda: False)
    with caplog.at_level("INFO", logger="telegram_bot"):
        assert telegram_bot._build_cs(None, None, None) == (None, None)
    assert "고객 문의(CS) 기능: 꺼짐" in caplog.text


def test_build_cs_on_returns_flow_and_job_and_logs_it(monkeypatch, tmp_path, caplog):
    from worker import cs_watch
    from worker.cs_flow import CsFlow

    monkeypatch.setattr(telegram_bot, "cs_enabled", lambda: True)
    monkeypatch.setattr(cs_watch, "default_state_dir", lambda: tmp_path)
    monkeypatch.setattr(cs_watch, "make_client", lambda: None)
    settings = type("S", (), {"allowed_chats": frozenset({1})})()
    with caplog.at_level("INFO", logger="telegram_bot"):
        flow, job = telegram_bot._build_cs(settings, object(), None)
    assert isinstance(flow, CsFlow) and callable(job) and "고객 문의(CS) 기능: 켜짐" in caplog.text
    assert "알림 시간 월화수목금 09:00-18:00" in caplog.text


# --- store document order -----------------------------------------------------------------------

def test_store_doc_comes_from_the_data_store_first(tmp_path, monkeypatch):
    local = tmp_path / "local.md"
    local.write_text("로컬 문서", encoding="utf-8")
    monkeypatch.setattr(cs_drafter, "doc_path", lambda: local)
    store = MemoryStore()
    store.write_text("cs/가게정보.md", "  저장소 문서  ", None, "m")
    assert load_store_doc(store=store) == "저장소 문서"


def test_store_doc_falls_back_to_the_local_file(tmp_path, monkeypatch):
    local = tmp_path / "local.md"
    local.write_text("로컬 문서", encoding="utf-8")
    monkeypatch.setattr(cs_drafter, "doc_path", lambda: local)
    store = MemoryStore()
    assert load_store_doc(store=store) == "로컬 문서"  # not in the store
    store.write_text("cs/가게정보.md", "  ", None, "m")
    assert load_store_doc(store=store) == "로컬 문서"  # empty in the store
    assert load_store_doc() == "로컬 문서"  # no store at all


def test_store_failure_falls_back_to_local_and_neither_is_the_old_error(tmp_path, monkeypatch):
    class Broken(MemoryStore):
        def read_text(self, path):
            raise RuntimeError("down")

    local = tmp_path / "local.md"
    monkeypatch.setattr(cs_drafter, "doc_path", lambda: local)
    with pytest.raises(DraftError, match="찾을 수 없습니다"):
        load_store_doc(store=Broken())
    local.write_text("로컬", encoding="utf-8")
    assert load_store_doc(store=Broken()) == "로컬"


def test_explicit_path_ignores_the_store(tmp_path):
    path = tmp_path / "x.md"
    path.write_text("지정 파일", encoding="utf-8")
    store = MemoryStore()
    store.write_text("cs/가게정보.md", "저장소", None, "m")
    assert load_store_doc(path, store=store) == "지정 파일"
