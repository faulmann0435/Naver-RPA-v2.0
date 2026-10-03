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
    assert "사장님" in store.read_text(OPTION_RULES_FILE).content
    last = store.history(OPTION_RULES_FILE)[0]
    assert last.author == "사장님" and last.message == "옵션규칙 수정: 수정 1, 추가 0, 삭제 0"
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
    assert store.history(SETTINGS_FILE)[0].author == "사장님"


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


def _store_with_workers(names: list[str]):
    store = _seeded_store()
    snapshot = store.read_text(SETTINGS_FILE)
    assert snapshot is not None
    data = {**json.loads(snapshot.content), "workers": names, "extra_key": 7}
    store.write_text(SETTINGS_FILE, json.dumps(data, ensure_ascii=False), snapshot.sha, "seed workers")
    return store


def _set_workers_edit(at: AppTest, added: list[str], deleted: list[int] | None = None) -> None:
    """A data_editor edit set through session_state only lasts for the next run."""
    sha = at.session_state["rules_data_settings"][1]
    at.session_state[f"rules_w_workers_{sha}"] = {
        "edited_rows": {}, "added_rows": [{"이름": n} for n in added], "deleted_rows": deleted or [],
    }


def _edit_workers(at: AppTest, added: list[str], deleted: list[int] | None = None) -> AppTest:
    _set_workers_edit(at, added, deleted)
    return at.run()


def test_settings_tab_saves_workers_and_keeps_other_keys():
    store = _store_with_workers(["갑", "을"])
    at = _run(_rules_page, store)
    _set_workers_edit(at, ["  병  "])
    at.button(key="rules_w_save_settings").click()
    at.run()
    assert not at.exception and not at.error
    data = json.loads(store.read_text(SETTINGS_FILE).content)
    assert data["workers"] == ["갑", "을", "병"] and data["extra_key"] == 7
    assert "작업자 목록" in store.history(SETTINGS_FILE)[0].message


def test_settings_tab_rejects_duplicate_and_empty_workers():
    store = _store_with_workers(["갑", "을"])
    sha = store.read_text(SETTINGS_FILE).sha
    at = _edit_workers(_run(_rules_page, store), ["갑"])
    assert any("겹칩니다" in e.value for e in at.error)
    assert at.button(key="rules_w_save_settings").disabled
    at = _edit_workers(_run(_rules_page, store), [], deleted=[0, 1])
    assert any("하나 이상" in e.value for e in at.error)
    assert store.read_text(SETTINGS_FILE).sha == sha


def _rules_page_with_sidebar() -> None:
    from ui import context, rules_page
    context.render_sidebar("테스트")
    rules_page.render()


def test_sidebar_picker_lists_workers_and_sets_commit_author():
    store = _store_with_workers(["갑", "을"])
    at = _run(_rules_page_with_sidebar, store)
    assert not at.exception
    picker = at.sidebar.selectbox(key="worker_name")
    assert picker.label == "작업자" and picker.options == ["갑", "을"] and picker.value == "갑"
    assert any("변경 이력에 남습니다" in c.value for c in at.sidebar.caption)
    picker.select("을").run()
    assert at.query_params["worker"] == ["을"]
    sha = store.read_text(SETTINGS_FILE).sha
    at.selectbox(key=f"rules_w_sep_{sha}").set_value(" + ").run()
    at.button(key="rules_w_save_settings").click()
    at.run()
    assert not at.exception and not at.error
    assert store.history(SETTINGS_FILE)[0].author == "을"


def test_sidebar_picker_restores_choice_from_query_param():
    at = AppTest.from_function(_rules_page_with_sidebar, default_timeout=TIMEOUT)
    at.session_state["_test_store"] = _store_with_workers(["갑", "을"])
    at.query_params["worker"] = "을"
    at.run()
    assert not at.exception and at.sidebar.selectbox(key="worker_name").value == "을"


def test_sidebar_without_store_uses_default_workers():
    at = AppTest.from_file("app.py", default_timeout=TIMEOUT).run()
    assert not at.exception
    assert at.sidebar.selectbox(key="worker_name").options == ["사장님", "사장님을노리는님", "개발자"]
