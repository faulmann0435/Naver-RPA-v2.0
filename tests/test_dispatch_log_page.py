"""송장 등록 기록 page: pure logic (labels, tables, filter) and an AppTest smoke run on a MemoryStore."""
from datetime import datetime

import pytest
from streamlit.testing.v1 import AppTest

from core.invoice_match import KST
from store import dispatch_queue as q
from store.memory_store import MemoryStore
from ui import dispatch_log_logic as logic

NOW = datetime(2026, 10, 13, 9, 5, tzinfo=KST)
ITEMS = [{"ids": ["1", "2"], "tracking": "111", "courier": "HANJIN", "recipient": "홍길동"},
         {"ids": ["3"], "tracking": "222", "courier": "CJGLS", "recipient": "김철수"},
         {"ids": ["4"], "tracking": "333", "courier": "NEWCODE", "recipient": "이영희"}]


def _request(status="done", source="telegram", **result):
    base = {"simulated": False, "success_ids": [], "failures": [], "error": None}
    return {"id": "r", "created_at": "2026-10-13T09:05:00+09:00", "by": "텔레그램 민수", "source": source,
            "status": status, "items": ITEMS, "result": None if status == "pending" else {**base, **result}}


@pytest.mark.parametrize("request_, label", [
    (_request("pending"), "⏳ 처리 대기"),
    (_request("failed", error="boom"), "❌ 실패"),
    (_request(simulated=True, success_ids=["1"]), "🧪 시험 모드 (실제 등록 안 됨)"),
    (_request(success_ids=["1"], failures=[{"id": "3", "code": "E", "message": "m"}]), "⚠️ 일부 실패"),
    (_request(success_ids=["1", "2", "3", "4"]), "✅ 완료"),
])
def test_status_label(request_, label):
    assert logic.status_label(request_) == label


def test_source_label_and_legacy_entry_without_source():
    assert logic.source_label(_request(source="telegram")) == "📱 텔레그램"
    legacy = {k: v for k, v in _request(source="app").items() if k != "source"}
    assert logic.source_label(legacy) == "🖥️ 사이트"


def test_filter_by_source():
    tg, app = _request(source="telegram"), _request(source="app")
    legacy = {k: v for k, v in app.items() if k != "source"}
    both = [tg, app, legacy]
    assert logic.filter_by_source(both, "전체") == both
    assert logic.filter_by_source(both, "텔레그램") == [tg]
    assert logic.filter_by_source(both, "사이트") == [app, legacy]


def test_summary_table_counts_and_kst_time():
    utc = {**_request(success_ids=["1", "2"], failures=[{"id": "3", "code": "E", "message": "m"}]),
           "created_at": "2026-10-13T00:05:00+00:00"}
    table = logic.summary_table([utc, _request("pending")])
    assert list(table.columns) == ["시간", "경로", "요청자", "상태", "송장 수", "성공", "실패"]
    assert table.iloc[0].to_dict() == {"시간": "10/13 09:05", "경로": "📱 텔레그램", "요청자": "텔레그램 민수",
                                       "상태": "⚠️ 일부 실패", "송장 수": 3, "성공": 2, "실패": 1}
    assert logic.summary_table([]).empty


def test_detail_table_results_and_courier_names():
    request = _request(success_ids=["1", "2"], failures=[{"id": "3", "code": "E", "message": "주소 오류"}])
    table = logic.detail_table(request)
    assert list(table["받는 사람"]) == ["홍길동", "김철수", "이영희"]
    assert list(table["택배사"]) == ["한진택배", "CJ대한통운", "NEWCODE"]  # unknown code shown as-is
    assert list(table["결과"]) == ["등록됨", "실패: 주소 오류", "–"]
    assert list(table["상품주문번호"]) == ["1, 2", "3", "4"]


def test_detail_table_simulated_pending_and_failed_run():
    assert list(logic.detail_table(_request(simulated=True, success_ids=["1", "2", "3", "4"]))["결과"]) == ["시험"] * 3
    assert set(logic.detail_table(_request("pending"))["결과"]) == {"대기"}
    failed = logic.detail_table(_request("failed", error="RuntimeError: x"))
    assert set(failed["결과"]) == {"실패: RuntimeError: x"}


def test_detail_table_of_legacy_entry_without_recipient():
    legacy = _request(success_ids=["1", "2"])
    legacy["items"] = [{"ids": ["1", "2"], "tracking": "111", "courier": "HANJIN"}]
    assert list(logic.detail_table(legacy)["받는 사람"]) == [""]


# --- the page -----------------------------------------------------------------------------------------

@pytest.fixture(autouse=True)
def _no_real_secrets(monkeypatch):
    from ui import context

    monkeypatch.setattr(context, "_data_store_settings", lambda: None)


def _page() -> None:
    from ui import dispatch_log_page
    dispatch_log_page.render()


def _run(store=None) -> AppTest:
    at = AppTest.from_function(_page, default_timeout=60)
    if store is not None:
        at.session_state["_test_store"] = store
    return at.run()


def test_page_without_store_shows_info():
    at = _run()
    assert not at.exception and any("데이터 저장소" in i.value for i in at.info)


def test_page_lists_both_sources_and_filters():
    store = MemoryStore()
    q.submit(store, [ITEMS[0]], "민수", NOW)
    q.record(store, "tg1", [ITEMS[1]], "텔레그램 민수", NOW, {"simulated": True, "success_ids": ["3"],
                                                          "failures": [], "error": None}, "done")
    at = _run(store)
    assert not at.exception and len(at.expander) == 2 and len(at.dataframe) == 3
    at.radio[0].set_value("텔레그램").run()
    assert not at.exception and len(at.expander) == 1 and "📱 텔레그램" in at.expander[0].label


def test_page_empty_log_and_store_error():
    at = _run(MemoryStore())
    assert any("기록이 없습니다" in i.value for i in at.info)

    class Broken(MemoryStore):
        def read_text(self, path):
            raise ValueError("broken")

    assert any("불러오지 못했습니다" in e.value for e in _run(Broken()).error)
