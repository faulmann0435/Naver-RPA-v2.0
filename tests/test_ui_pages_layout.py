"""AppTest smoke + flow tests of 발주서 양식 and the layout history (in-memory store, no network)."""
import dataclasses

import pytest
from streamlit.testing.v1 import AppTest

from core.config_loader import read_config_sheets
from store.csv_codec import from_csv_text
from store.layout_repo import OUTPUT_LAYOUT_FILE, layout_sheet_to_csv
from store.rules_repo import load_config_from_store
from tests.regression_harness import CONFIG_PATH
from tests.test_ui_pages import TIMEOUT, _seeded_store

FORM = "메로 발주양식"
UNUSED = "임시 양식"


@pytest.fixture(autouse=True)
def _no_real_secrets(monkeypatch):
    from ui import context

    monkeypatch.setattr(context, "_data_store_settings", lambda: None)


def _layout_page() -> None:
    from ui import layout_page
    layout_page.render()


def _history_page() -> None:
    from ui import history_page
    history_page.render()


def _sidebar() -> None:
    from ui import context
    context.render_sidebar("테스트 저장소")


def _layout_store(migrated: bool = True, extra: str = ""):
    """extra: raw CSV lines appended to the layout (every real form is used by a rule, so tests add their own)."""
    store = _seeded_store()
    if migrated:
        sheet = read_config_sheets(str(CONFIG_PATH))["OutputLayout"]
        text = layout_sheet_to_csv(sheet, "migration", "2026-10-02T00:00:00+09:00") + extra
        store.write_text(OUTPUT_LAYOUT_FILE, text, None, "seed layout")
    return store


def _run(func, store=None) -> AppTest:
    at = AppTest.from_function(func, default_timeout=TIMEOUT)
    if store is not None:
        at.session_state["_test_store"] = store
    return at.run()


def _select_form(at: AppTest, form: str) -> AppTest:
    at.selectbox(key="layout_w_form").set_value(form).run()
    return at


def _set_edit(at: AppTest, form: str, edited=None, added=None) -> None:
    """A data_editor edit set through session_state only lasts for the next run."""
    key = f"layout_w_grid_{at.session_state['layout_ver']}_{form}"
    at.session_state[key] = {"edited_rows": edited or {}, "added_rows": added or [], "deleted_rows": []}


def _grid_edit(at: AppTest, form: str, edited=None, added=None) -> AppTest:
    _set_edit(at, form, edited, added)
    return at.run()


def _edit_and_save(at: AppTest, form: str, edited=None, added=None) -> AppTest:
    """Edit and click 저장 in the same run (the edit does not survive into a later run)."""
    _set_edit(at, form, edited, added)
    at.button(key="layout_w_save").click()
    return at.run()


def test_page_without_store_shows_info():
    at = _run(_layout_page)
    assert not at.exception and any("데이터 저장소가 설정되지 않아" in i.value for i in at.info)


def test_not_migrated_shows_migration_hint_and_read_only_table():
    at = _run(_layout_page, _layout_store(migrated=False))
    assert not at.exception
    assert any("이관 도구" in i.value for i in at.info) and len(at.dataframe) == 1
    assert len(at.button) == 1  # only the reload button: nothing to save


def test_page_renders_form_grid_preview_and_help():
    at = _run(_layout_page, _layout_store())
    assert not at.exception and not at.error
    assert at.selectbox(key="layout_w_form").value == "프리미엄과메기"
    assert any("왼쪽부터 몇 번째 칸" in c.value for c in at.caption)
    assert any("미리보기" in m.value for m in at.markdown) and len(at.dataframe) >= 1
    assert at.button(key="layout_w_save").disabled


def test_sidebar_shows_where_the_layout_comes_from():
    migrated = _run(_sidebar, _layout_store())
    assert any(c.value == "양식 출처: 데이터 저장소" for c in migrated.sidebar.caption)
    fallback = _run(_sidebar, _layout_store(migrated=False))
    assert any(c.value == "양식 출처: config.xlsx" for c in fallback.sidebar.caption)


def test_edit_and_save_writes_store_with_author_and_message():
    store = _layout_store()
    at = _select_form(_run(_layout_page, store), FORM)
    before = store.read_text(OUTPUT_LAYOUT_FILE).content
    assert not _grid_edit(at, FORM, edited={1: {"칸 이름": "새이름"}}).button(key="layout_w_save").disabled
    _edit_and_save(at, FORM, edited={1: {"칸 이름": "새이름"}})
    assert not at.exception and not at.error
    saved = store.read_text(OUTPUT_LAYOUT_FILE).content
    assert saved != before and "새이름" in saved and "사장님" in saved
    last = store.history(OUTPUT_LAYOUT_FILE)[0]
    assert last.author == "사장님" and last.message == "발주서 양식 수정: 메로 발주양식 (수정 1, 추가 0, 삭제 0)"
    assert any("저장했습니다" in s.value for s in at.success)
    assert at.session_state["layout_data"].sha == store.read_text(OUTPUT_LAYOUT_FILE).sha  # reloaded
    config = load_config_from_store(store, "config.xlsx", "1111")
    assert "새이름" in set(config["OutputLayout"]["HeaderName"])


def test_stale_sha_shows_conflict_and_does_not_write():
    store = _layout_store()
    at = _select_form(_run(_layout_page, store), FORM)
    at.session_state["layout_data"] = dataclasses.replace(at.session_state["layout_data"], sha="stale")
    sha = store.read_text(OUTPUT_LAYOUT_FILE).sha
    _edit_and_save(at, FORM, edited={1: {"칸 이름": "새이름"}})
    assert store.read_text(OUTPUT_LAYOUT_FILE).sha == sha
    assert any("다른 사람이 먼저 저장했습니다" in e.value for e in at.error)


def test_duplicate_header_blocks_save():
    store = _layout_store()
    at = _select_form(_run(_layout_page, store), FORM)
    sha = store.read_text(OUTPUT_LAYOUT_FILE).sha
    assert any("중복" in e.value for e in _grid_edit(at, FORM, edited={1: {"칸 이름": "수취인명"}}).error)
    _edit_and_save(at, FORM, edited={1: {"칸 이름": "수취인명"}})
    assert store.read_text(OUTPUT_LAYOUT_FILE).sha == sha and any("오류가 있어" in e.value for e in at.error)


def test_warning_needs_confirmation_before_saving():
    store = _layout_store()
    at = _select_form(_run(_layout_page, store), FORM)
    sha = store.read_text(OUTPUT_LAYOUT_FILE).sha
    edit = {0: {"고정 글자": "고정"}}  # row 0 already has a source: both set = warning
    _edit_and_save(at, FORM, edited=edit)
    assert store.read_text(OUTPUT_LAYOUT_FILE).sha == sha and any("경고를 확인" in w.value for w in at.warning)
    at.checkbox(key="layout_w_confirm").check()
    _edit_and_save(at, FORM, edited=edit)
    assert store.read_text(OUTPUT_LAYOUT_FILE).sha != sha


def test_add_form_and_save():
    store = _layout_store()
    at = _run(_layout_page, store)
    at.text_input(key="layout_w_new_name").set_value("새 양식")
    at.text_input(key="layout_w_new_file").set_value("새파일")
    at.button(key="layout_w_new_add").click().run()
    assert not at.exception and at.selectbox(key="layout_w_form").value == "새 양식"
    assert any("칸이 하나도 없습니다" in e.value for e in at.error) and at.button(key="layout_w_save").disabled
    _edit_and_save(at, "새 양식", added=[{"칸 이름": "받는분", "넣을 내용": "받는 분 이름"}])
    assert not at.exception
    new = from_csv_text(store.read_text(OUTPUT_LAYOUT_FILE).content, as_text=True)
    rows = new[new["양식명칭"] == "새 양식"]
    assert rows["파일명"].tolist() == ["새파일"] and rows["매핑데이터"].tolist() == ["수취인명"]
    assert at.selectbox(key="layout_w_form").value == "새 양식"


def test_add_form_rejects_existing_name():
    at = _run(_layout_page, _layout_store())
    at.text_input(key="layout_w_new_name").set_value(FORM)
    at.button(key="layout_w_new_add").click().run()
    assert any("이미 있습니다" in e.value for e in at.error)


def test_delete_referenced_form_is_blocked_unreferenced_is_deleted():
    store = _layout_store(extra="임시 양식,임시파일,A,이름,수취인명,,naver,2026-10-02T00:00:00+09:00,migration\n")
    at = _select_form(_run(_layout_page, store), "속초 발주양식")
    sha = store.read_text(OUTPUT_LAYOUT_FILE).sha
    at.checkbox(key="layout_w_delete_confirm").check().run()
    at.button(key="layout_w_delete").click().run()
    assert store.read_text(OUTPUT_LAYOUT_FILE).sha == sha and any("아직 쓰이고 있어" in e.value for e in at.error)
    at = _select_form(_run(_layout_page, store), UNUSED)
    at.checkbox(key="layout_w_delete_confirm").check().run()
    at.button(key="layout_w_delete").click().run()
    assert not at.exception
    assert UNUSED not in store.read_text(OUTPUT_LAYOUT_FILE).content
    assert store.history(OUTPUT_LAYOUT_FILE)[0].message == f"발주서 양식 삭제: {UNUSED}"


def test_layout_history_diff_and_revert():
    store = _layout_store()
    at = _select_form(_run(_layout_page, store), FORM)
    original = store.read_text(OUTPUT_LAYOUT_FILE).content
    _edit_and_save(at, FORM, edited={1: {"칸 이름": "새이름"}})
    assert store.read_text(OUTPUT_LAYOUT_FILE).content != original
    page = _run(_history_page, store)
    page.selectbox(key="history_target").set_value("발주서 양식").run()
    page.selectbox(key=f"history_pick_{OUTPUT_LAYOUT_FILE}").set_value(1).run()
    assert not page.exception and not page.error
    assert any("새이름" in str(df.value.to_dict()) for df in page.dataframe)
    page.checkbox(key=f"history_confirm_{OUTPUT_LAYOUT_FILE}").check().run()
    page.button(key="history_revert").click().run()
    assert not page.exception and not page.error
    assert store.read_text(OUTPUT_LAYOUT_FILE).content == original
    assert store.history(OUTPUT_LAYOUT_FILE)[0].message.startswith("되돌리기: 발주서 양식 → ")
