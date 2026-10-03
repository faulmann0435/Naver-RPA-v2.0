"""Page smoke tests with streamlit's AppTest, using only an in-memory store (no network)."""
import dataclasses

import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

from store.csv_codec import to_csv_text
from store.dictionary_repo import group_id, load_dictionary_frame
from store.memory_store import MemoryStore
from store.rules_repo import DICTIONARY_COLUMNS, DICTIONARY_FILE
from tests.test_ui_logic import make_store
from ui.dictionary_logic import DELETE
from ui.option_form import state_key
from ui.order_logic import item_prefix, unmatched_items
from ui.product_logic import option_prefix
from ui.product_page import vendor_keys

TIMEOUT = 60


@pytest.fixture(autouse=True)
def _no_real_secrets(monkeypatch):
    """Never touch a real data store: only an injected MemoryStore may exist."""
    from ui import context

    monkeypatch.setattr(context, "_data_store_settings", lambda: None)


def _order_page() -> None:
    from ui import order_page
    order_page.render()


def _dictionary_page() -> None:
    from ui import dictionary_page
    dictionary_page.render()


def _preview_page() -> None:
    from ui import preview_page
    preview_page.render()


GROUP_APPLE, GROUP_PEAR = group_id("100", "사과"), group_id("101", "배")


def _seeded_store() -> MemoryStore:
    store = make_store()
    rows = [
        {"channel": "naver", "enabled": "1", "product_no": str(100 + i), "option_key": "", "product_name_ref": n,
         "vendor_id": "메로 발주 양식".replace(" ", "") if False else "속초 발주양식",
         "display_template": f"{n} {{수량}}개", "needs_review": "1"}
        for i, n in enumerate(["사과", "배"])
    ]
    frame = pd.DataFrame([{**dict.fromkeys(DICTIONARY_COLUMNS, ""), **r} for r in rows], columns=DICTIONARY_COLUMNS)
    store.write_text(DICTIONARY_FILE, to_csv_text(frame), load_dictionary_frame(store)[1], "seed dict")
    return store


def _run(func, store: MemoryStore | None = None) -> AppTest:
    at = AppTest.from_function(func, default_timeout=TIMEOUT)
    if store is not None:
        at.session_state["_test_store"] = store
    return at.run()


@pytest.mark.parametrize("func", [_order_page, _dictionary_page, _preview_page])
def test_pages_render_without_exception(func):
    at = _run(func, _seeded_store())
    assert not at.exception


def test_dictionary_page_without_store_shows_info():
    at = _run(_dictionary_page)
    assert not at.exception and any("데이터 저장소" in i.value for i in at.info)


def test_full_app_navigates_all_pages():
    at = AppTest.from_file("app.py", default_timeout=TIMEOUT)
    at.session_state["_test_store"] = _seeded_store()
    at.run()
    assert not at.exception
    for page in ("dictionary", "preview", "order", "rules", "history"):
        at.switch_page(f"ui/{page}_page.py").run()
        assert not at.exception


def _edit_and_save(store: MemoryStore, mutate, base_sha: str | None = None) -> AppTest:
    at = _run(_dictionary_page, store)
    base, _sha = at.session_state["dict_base"]
    work = at.session_state["dict_work"].copy()
    mutate(work)
    at.session_state["dict_work"] = work
    if base_sha is not None:
        at.session_state["dict_base"] = (base, base_sha)
    at.run()
    next(b for b in at.button if b.label == "저장").click()
    return at.run()


def test_save_changes_store_and_records_author():
    store = _seeded_store()
    before = load_dictionary_frame(store)[1]
    at = _edit_and_save(store, lambda w: w.__setitem__("display_template", ["사과 상자 {수량}", "배 {수량}개"]))
    assert not at.exception and not at.error
    frame, sha = load_dictionary_frame(store)
    assert sha != before and frame.loc[0, "display_template"] == "사과 상자 {수량}"
    assert frame.loc[0, "updated_by"] == "사장님" and frame.loc[1, "updated_by"] == ""
    assert store.history(DICTIONARY_FILE)[0].author == "사장님"
    assert store.history(DICTIONARY_FILE)[0].message == "사전 수정: 수정 1, 삭제 0"


def test_filter_edit_save_keeps_hidden_rows():
    store = _seeded_store()
    at = _run(_dictionary_page, store)
    at.text_input(key="dict_query").set_value("사과").run()
    work = at.session_state["dict_work"].copy()
    work.loc[0, "display_template"] = "사과 {수량}박스"  # what the folded editor edit produces
    at.session_state["dict_work"] = work
    at.run()
    next(b for b in at.button if b.label == "저장").click()
    at.run()
    frame, _ = load_dictionary_frame(store)
    assert list(frame["product_name_ref"]) == ["사과", "배"]
    assert frame.loc[0, "display_template"] == "사과 {수량}박스"


def test_delete_flag_removes_row_on_save():
    store = _seeded_store()
    _edit_and_save(store, lambda w: w.__setitem__(DELETE, [False, True]))
    assert list(load_dictionary_frame(store)[0]["product_name_ref"]) == ["사과"]


def test_invalid_edit_blocks_save():
    store = _seeded_store()
    sha = load_dictionary_frame(store)[1]
    at = _edit_and_save(store, lambda w: w.__setitem__("vendor_id", ["없는양식", "속초 발주양식"]))
    assert load_dictionary_frame(store)[1] == sha
    assert any("발주양식" in e.value for e in at.error)


def test_stale_sha_shows_conflict_error():
    store = _seeded_store()
    sha = load_dictionary_frame(store)[1]
    at = _edit_and_save(store, lambda w: w.__setitem__("display_template", ["X {수량}", "Y {수량}"]), base_sha="stale")
    assert load_dictionary_frame(store)[1] == sha
    assert any("다른 사람이 먼저 저장했습니다" in e.value for e in at.error)


def test_order_page_shows_results_from_session_state():
    from core.pipeline import process_orders
    from tests.test_ui_logic import ORDER
    from ui.context import vendor_ids
    from ui.order_logic import build_suggestions

    store = make_store()
    at = AppTest.from_function(_order_page, default_timeout=TIMEOUT)
    at.session_state["_test_store"] = store
    at.run()
    from core.config_loader import read_config_sheets  # noqa: F401
    from store.rules_repo import load_config_from_store
    config = load_config_from_store(store, "config.xlsx", "1111")
    result = process_orders(ORDER, config)
    result = dataclasses.replace(result, stats={**result.stats, "rows_unclassified": 1})
    at.session_state["order_result"] = result
    at.session_state["order_df"] = ORDER
    at.session_state["order_suggestions"] = build_suggestions(result.unmatched, config, vendor_ids(config))
    at.run()
    assert not at.exception
    assert any("사전 미등록 2건" in w.value for w in at.warning)
    assert any("미분류" in w.value and "들어가지 않았습니다" in w.value for w in at.warning)
    assert [m.label for m in at.metric] == ["사전 적용", "자동 규칙", "미분류", "검토 필요 사전 항목"]
    assert any(b.label == "선택한 항목 사전에 등록하고 다시 처리" for b in at.button)
    items = unmatched_items(at.session_state["order_suggestions"])
    assert len(items) == 2
    for item in items:
        prefix = item_prefix(item)
        assert at.checkbox(key=state_key(prefix, "register")).value is False
        assert at.text_input(key=state_key(prefix, "sentence")).label == "발주서에 적을 문장"


def test_order_page_registers_checked_item_with_built_template():
    from core.pipeline import process_orders
    from store.rules_repo import load_config_from_store
    from tests.test_ui_logic import ORDER
    from ui.context import vendor_ids
    from ui.order_logic import build_suggestions

    store = make_store()
    config = load_config_from_store(store, "config.xlsx", "1111")
    result = process_orders(ORDER, config)
    at = AppTest.from_function(_order_page, default_timeout=TIMEOUT)
    at.session_state["_test_store"] = store
    at.run()
    at.session_state["order_result"] = result
    at.session_state["order_df"] = ORDER
    suggestions = build_suggestions(result.unmatched, config, vendor_ids(config))
    at.session_state["order_suggestions"] = suggestions
    at.run()
    item = next(i for i in unmatched_items(suggestions) if i.name == "메로 구이")
    prefix = item_prefix(item)
    at.text_input(key=state_key(prefix, "sentence")).set_value("메로 구이 3마리").run()
    at.checkbox(key=state_key(prefix, "scale_0")).check().run()
    at.checkbox(key=state_key(prefix, "register")).check().run()
    next(b for b in at.button if b.label == "선택한 항목 사전에 등록하고 다시 처리").click().run()
    assert not at.exception
    frame, _ = load_dictionary_frame(store)
    assert list(frame["display_template"]) == ["메로 구이 {수량*3}마리"]
    assert frame.loc[0, "updated_by"] == "사장님"


def test_preview_page_runs_a_bundle():
    at = _run(_preview_page, _seeded_store())
    at.session_state["preview_rows"] = pd.DataFrame(
        {"상품번호": ["100"], "상품명": ["사과"], "옵션정보": [""], "수량": [3]}
    )
    at.run()
    next(b for b in at.button if b.label == "결과 보기").click()
    at.run()
    assert not at.exception and at.dataframe


def test_bulk_uncheck_review_on_filtered_rows_then_save():
    store = _seeded_store()
    at = _run(_dictionary_page, store)
    at.text_input(key="dict_query").set_value("사과").run()        # filter: only 사과 visible
    at.selectbox(key="dict_bulk_target").set_value("검토 필요").run()
    next(b for b in at.button if b.label == "모두 해제").click().run()
    assert not at.exception
    assert at.session_state["dict_work"]["needs_review"].tolist() == ["0", "1"]  # 배 (hidden) untouched
    next(b for b in at.button if b.label == "저장").click().run()
    frame, _sha = load_dictionary_frame(store)
    assert frame["needs_review"].tolist() == ["0", "1"]


# ---------------------------------------------------------------- product-by-product tab

def _product_at(store: MemoryStore, selected: str | None = GROUP_APPLE) -> AppTest:
    at = AppTest.from_function(_dictionary_page, default_timeout=TIMEOUT)
    at.session_state["_test_store"] = store
    if selected is not None:
        at.session_state["pm_selected"] = selected
    return at.run()


def test_product_tab_renders_two_tabs_and_list():
    at = _product_at(_seeded_store(), selected=None)
    assert not at.exception
    assert [t.label for t in at.tabs] == ["상품별 편집", "표로 한꺼번에 보기"]
    assert any("왼쪽 목록에서 상품을 고르세요" in i.value for i in at.info)
    assert any("수량에 따라 늘어나는 숫자" in c.value for c in at.caption)


def test_product_tab_shows_option_form_for_selected_product():
    at = _product_at(_seeded_store())
    prefix = option_prefix("100", "")
    assert not at.exception
    assert at.text_input(key=state_key(prefix, "sentence")).value == "사과 1개"
    assert at.checkbox(key=state_key(prefix, "scale_0")).value is True
    assert [t.value for t in at.text if t.value.startswith(("1개", "2개", "3개"))] == [
        "1개  →  사과 1개", "2개  →  사과 2개", "3개  →  사과 3개"
    ]
    assert any(b.label == "이 상품 저장" for b in at.button)


def test_product_save_builds_template_and_records_author():
    store = _seeded_store()
    at = _product_at(store)
    prefix = option_prefix("100", "")
    at.text_input(key=state_key(prefix, "sentence")).set_value("사과 8개 박스 1개").run()
    assert [at.checkbox(key=state_key(prefix, f"scale_{i}")).value for i in (0, 1)] == [False, True]
    at.checkbox(key=state_key(prefix, "scale_0")).check().run()
    at.checkbox(key=state_key(prefix, "done")).check().run()
    at.button(key="pm_save").click().run()
    assert not at.exception and not at.error
    frame, _sha = load_dictionary_frame(store)
    assert frame.loc[0, "display_template"] == "사과 {수량*8}개 박스 {수량}개"
    assert frame.loc[0, "needs_review"] == "0" and frame.loc[1, "display_template"] == "배 {수량}개"
    assert frame.loc[0, "updated_by"] == "사장님" and frame.loc[1, "updated_by"] == ""
    last = store.history(DICTIONARY_FILE)[0]
    assert last.author == "사장님" and last.message == "품목 수정: 사과 (수정 1, 삭제 0)"
    assert any("저장했습니다" in s.value for s in at.success)
    assert at.text_input(key=state_key(prefix, "sentence")).value == "사과 8개 박스 1개"


def test_product_pending_edits_survive_switching_products():
    store = _seeded_store()
    at = _product_at(store)
    prefix = option_prefix("100", "")
    at.text_input(key=state_key(prefix, "sentence")).set_value("사과 5개").run()
    at.checkbox(key=state_key(prefix, "scale_0")).uncheck().run()
    at.session_state["pm_selected"] = GROUP_PEAR
    at.run()
    assert not at.exception and at.text_input(key=state_key(option_prefix("101", ""), "sentence")).value == "배 1개"
    at.session_state["pm_selected"] = GROUP_APPLE
    at.run()
    assert at.text_input(key=state_key(prefix, "sentence")).value == "사과 5개"
    assert at.checkbox(key=state_key(prefix, "scale_0")).value is False
    assert load_dictionary_frame(store)[0].loc[0, "display_template"] == "사과 {수량}개"
    at.button(key="pm_cancel").click().run()
    assert at.text_input(key=state_key(prefix, "sentence")).value == "사과 1개"


def test_product_delete_option_and_blocked_invalid_save():
    store = _seeded_store()
    at = _product_at(store)
    prefix = option_prefix("100", "")
    at.text_input(key=state_key(prefix, "sentence")).set_value("").run()
    at.button(key="pm_save").click().run()
    assert any("발주서 표기" in e.value for e in at.error)
    assert load_dictionary_frame(store)[0].loc[0, "display_template"] == "사과 {수량}개"
    at.text_input(key=state_key(prefix, "sentence")).set_value("사과 {수량}개").run()
    at.checkbox(key=state_key(prefix, "delete")).check().run()
    at.button(key="pm_save").click().run()
    assert list(load_dictionary_frame(store)[0]["product_name_ref"]) == ["배"]


def test_product_vendor_change_applies_to_options():
    store = _seeded_store()
    at = _product_at(store)
    vendors = list(at.selectbox(key=vendor_keys(GROUP_APPLE)[0]).options)
    other = next(v for v in vendors if v != "속초 발주양식")
    at.selectbox(key=vendor_keys(GROUP_APPLE)[0]).set_value(other).run()
    assert at.selectbox(key=state_key(option_prefix("100", ""), "vendor")).value == other
    at.button(key="pm_save").click().run()
    assert load_dictionary_frame(store)[0].loc[0, "vendor_id"] == other


def _addon_store() -> MemoryStore:
    store = make_store()
    rows = [
        {"product_no": "200", "option_key": "a", "product_name_ref": "코다리", "display_template": "코다리 {수량}개"},
        {"product_no": "200", "option_key": "x", "product_name_ref": "명태회무침", "display_template": "회무침 {수량}개"},
    ]
    base = {**dict.fromkeys(DICTIONARY_COLUMNS, ""), "channel": "naver", "enabled": "1", "vendor_id": "속초 발주양식", "needs_review": "1"}
    frame = pd.DataFrame([{**base, **r} for r in rows], columns=DICTIONARY_COLUMNS)
    store.write_text(DICTIONARY_FILE, to_csv_text(frame), load_dictionary_frame(store)[1], "seed addon")
    return store


def test_product_tab_lists_main_and_addon_as_separate_groups_and_saves_one():
    store = _addon_store()
    addon = group_id("200", "명태회무침")
    at = _product_at(store, selected=addon)
    assert not at.exception
    assert any(m.value == "### 명태회무침" for m in at.markdown)
    assert any("상품 2개" in c.value for c in at.caption)
    prefix = option_prefix("200", "x")
    at.text_input(key=state_key(prefix, "sentence")).set_value("회무침 3개").run()
    at.checkbox(key=state_key(prefix, "scale_0")).check().run()
    at.button(key="pm_save").click().run()
    assert not at.exception and not at.error
    frame, _sha = load_dictionary_frame(store)
    assert list(frame["product_name_ref"]) == ["코다리", "명태회무침"]
    assert frame.loc[0, "display_template"] == "코다리 {수량}개" and frame.loc[1, "display_template"] != "회무침 {수량}개"
    assert store.history(DICTIONARY_FILE)[0].message.startswith("품목 수정: 명태회무침")
