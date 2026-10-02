"""AppTest smoke + flow tests of 고급 설정, 변경 이력 and the Excel import (in-memory store, no network)."""
import dataclasses
import json

import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

from store.csv_codec import to_csv_text
from store.rules_repo import (
    DICTIONARY_COLUMNS,
    DICTIONARY_FILE,
    OPTION_RULES_FILE,
    PRODUCT_ROUTE_FILE,
    SETTINGS_FILE,
    load_config_from_store,
)
from tests.test_ui_pages import TIMEOUT, _seeded_store
from ui.rules_logic import with_work

PARAM, TARGET, NOTE = "Parameter (설정값)", "적용대상(상품명)", "설명 (비고)"


@pytest.fixture(autouse=True)
def _no_real_secrets(monkeypatch):
    from ui import context

    monkeypatch.setattr(context, "_data_store_settings", lambda: None)


def _rules_page() -> None:
    from ui import rules_page
    rules_page.render()


def _history_page() -> None:
    from ui import history_page
    history_page.render()


def _dictionary_page() -> None:
    from ui import dictionary_page
    dictionary_page.render()


def _run(func, store) -> AppTest:
    at = AppTest.from_function(func, default_timeout=TIMEOUT)
    at.session_state["_test_store"] = store
    return at.run()


def _edit(at: AppTest, key: str, column: str, first_value: str) -> None:
    state = at.session_state[key]
    work = state.work.copy()
    work[column] = [first_value, *list(work[column][1:])]
    at.session_state[key] = with_work(state, work)
    at.run()


def _save(at: AppTest, kind: str) -> None:
    at.button(key=f"rules_w_save_{kind}").click()
    at.run()


def test_new_pages_render_without_exception():
    for func in (_rules_page, _history_page):
        assert not _run(func, _seeded_store()).exception


def test_rules_page_tabs_warning_and_help():
    at = _run(_rules_page, _seeded_store())
    assert [t.label for t in at.tabs] == ["상품분류", "옵션규칙", "규칙 시험", "설정"]
    assert any("일반 작업은 '품목 관리'에서 하세요" in w.value for w in at.warning)
    assert any(e.label == "ActionType 설명" for e in at.expander)


def test_rules_save_changes_store_and_config_still_parses():
    store = _seeded_store()
    at = _run(_rules_page, store)
    _edit(at, "rules_data_options", PARAM, "팩")
    _save(at, "options")
    assert not at.exception and not at.error
    assert "faulmann0435@gmail.com" in store.read_text(OPTION_RULES_FILE).content
    last = store.history(OPTION_RULES_FILE)[0]
    assert last.author == "faulmann0435@gmail.com" and last.message == "옵션규칙 수정: 수정 1, 추가 0, 삭제 0"
    config = load_config_from_store(store, "config.xlsx", "1111")
    assert config["OptionRules"]["Parameter"].iloc[0] == "팩" and len(config["OptionRules"]) == 60


def test_removed_default_row_blocks_save():
    store = _seeded_store()
    sha = store.read_text(PRODUCT_ROUTE_FILE).sha
    at = _run(_rules_page, store)
    state = at.session_state["rules_data_route"]
    work = state.work.copy()
    work["_delete"] = [k == "DEFAULT" for k in work["키워드"]]
    at.session_state["rules_data_route"] = with_work(state, work)
    at.run()
    _save(at, "route")
    assert store.read_text(PRODUCT_ROUTE_FILE).sha == sha and any("DEFAULT" in e.value for e in at.error)


def test_warning_needs_confirmation_before_saving():
    store = _seeded_store()
    sha = store.read_text(OPTION_RULES_FILE).sha
    at = _run(_rules_page, store)
    _edit(at, "rules_data_options", TARGET, "문어, 오징어")
    _save(at, "options")
    assert store.read_text(OPTION_RULES_FILE).sha == sha and any("경고를 확인" in w.value for w in at.warning)
    at.checkbox(key="rules_w_confirm_options").check().run()
    _save(at, "options")
    assert store.read_text(OPTION_RULES_FILE).sha != sha


def test_stale_sha_shows_conflict_message():
    store = _seeded_store()
    at = _run(_rules_page, store)
    at.session_state["rules_data_options"] = dataclasses.replace(at.session_state["rules_data_options"], sha="stale")
    _edit(at, "rules_data_options", NOTE, "메모")
    sha = store.read_text(OPTION_RULES_FILE).sha
    _save(at, "options")
    assert store.read_text(OPTION_RULES_FILE).sha == sha
    assert any("다른 사람이 먼저 저장했습니다" in e.value for e in at.error)


def test_settings_save_writes_settings_json():
    store = _seeded_store()
    at = _run(_rules_page, store)
    sha = store.read_text(SETTINGS_FILE).sha
    at.selectbox(key=f"rules_w_sep_{sha}").set_value(" + ").run()
    at.button(key="rules_w_save_settings").click()
    at.run()
    assert not at.exception and not at.error
    data = json.loads(store.read_text(SETTINGS_FILE).content)
    assert data["item_separator"] == " + " and data["ignored_option_groups"]
    assert store.history(SETTINGS_FILE)[0].author == "faulmann0435@gmail.com"


def test_rule_test_tab_shows_vendor_and_final_text():
    at = _run(_rules_page, _seeded_store())
    at.text_input(key="rules_w_t_name").set_value("메로 구이")
    at.text_input(key="rules_w_t_option").set_value("메로 선택: 반마리")
    at.run()
    assert not at.exception
    assert any("분류된 발주양식" in m.value for m in at.markdown) and any("최종 문구" in m.value for m in at.markdown)


def test_revert_flow_on_memory_store():
    store = _seeded_store()
    at = _run(_rules_page, store)
    _edit(at, "rules_data_options", PARAM, "팩")
    _save(at, "options")
    edited = store.read_text(OPTION_RULES_FILE).content
    seed = next(r for r in store.history(OPTION_RULES_FILE) if r.message == "seed")
    original = store.read_text_at(OPTION_RULES_FILE, seed.sha)
    page = _run(_history_page, store)
    page.selectbox(key="history_target").set_value("옵션규칙").run()
    page.selectbox(key=f"history_pick_{OPTION_RULES_FILE}").set_value(1).run()
    assert not page.exception
    page.checkbox(key=f"history_confirm_{OPTION_RULES_FILE}").check().run()
    page.button(key="history_revert").click().run()
    assert not page.exception and not page.error
    assert store.read_text(OPTION_RULES_FILE).content == original != edited
    assert store.history(OPTION_RULES_FILE)[0].message.startswith("되돌리기: 옵션규칙 → ")


def test_revert_blocked_when_old_dictionary_is_invalid():
    store = _seeded_store()
    sha = store.read_text(DICTIONARY_FILE).sha
    row = {**dict.fromkeys(DICTIONARY_COLUMNS, ""), "channel": "naver", "enabled": "1", "product_no": "1",
           "vendor_id": "없는양식", "display_template": "x"}
    sha = store.write_text(DICTIONARY_FILE, to_csv_text(pd.DataFrame([row])), sha, "나쁜 버전")
    store.write_text(DICTIONARY_FILE, to_csv_text(pd.DataFrame(columns=DICTIONARY_COLUMNS)), sha, "비움")
    at = _run(_history_page, store)
    at.selectbox(key="history_target").set_value("품목 사전").run()
    at.selectbox(key=f"history_pick_{DICTIONARY_FILE}").set_value(1).run()
    assert any("되돌릴 수 없습니다" in e.value for e in at.error)
    assert at.button(key="history_revert").disabled


def test_dictionary_grid_tab_has_excel_import():
    at = _run(_dictionary_page, _seeded_store())
    assert not at.exception and any("엑셀 가져오기" in m.value for m in at.markdown)
