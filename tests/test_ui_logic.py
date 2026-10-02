"""Pure UI helpers: order suggestions/registration, dictionary view/fold/save, bundle preview."""
import pandas as pd
import pytest

from core.config_loader import read_config_sheets
from core.dictionary import DictionarySettings, ItemDictionary
from core.pipeline import process_orders
from store.base import Author
from store.dictionary_repo import load_dictionary_frame, load_state
from store.memory_store import MemoryStore
from store.rules_repo import (
    DICTIONARY_COLUMNS,
    load_config_from_store,
    sheets_to_rule_csvs,
)
from store.validators import validate_dictionary
from ui.context import vendor_ids
from ui.dictionary_logic import (
    DELETE,
    RID,
    V_DELETE,
    V_TEMPLATE,
    build_view,
    change_table,
    changed_rids,
    describe_issue,
    filter_rids,
    fold_edits,
    kept_rows,
    make_work,
    prepare_save,
)
from ui.order_logic import (
    COL_CHECK,
    COL_TEMPLATE,
    COL_VENDOR,
    build_suggestions,
    register_rows,
    selection_preview,
)
from ui.preview_logic import preview_bundle

SETTINGS = DictionarySettings()
USER = Author(name="u@example.com", email="u@example.com")
ORDER = pd.DataFrame({
    "상품주문번호": ["A1", "A2"], "상품번호": ["111", "222"], "상품명": ["메로 구이", "알수없는 상품"],
    "옵션정보": ["메로 선택: 반마리", "기타: 하나"], "수량": [2, 1],
    "배송비 묶음번호": ["B1", "B2"], "결제일": ["2026-01-01", "2026-01-02"],
})


def make_store() -> MemoryStore:
    store = MemoryStore()
    files = sheets_to_rule_csvs(read_config_sheets("config.xlsx", "1111"), "test", "2026-01-01")
    for name, text in files.items():
        store.write_text(name, text, None, "seed")
    return store


@pytest.fixture(scope="module")
def config() -> dict:
    return load_config_from_store(make_store(), "config.xlsx", "1111")


def test_suggestions_register_and_reprocess(config):
    store = make_store()
    ids = vendor_ids(config)
    dictionary, settings, _ = load_state(store)
    first = process_orders(ORDER, config, dictionary, settings)
    assert first.stats["rows_dictionary"] == 0 and len(first.unmatched) == 2
    sug = build_suggestions(first.unmatched, config, ids, settings)
    assert set(sug[COL_VENDOR]) <= set(ids) and not sug[COL_CHECK].any()
    meroe = sug[sug["상품명"] == "메로 구이"].index[0]
    edited = sug.copy()
    edited.loc[meroe, COL_CHECK] = True
    edited.loc[meroe, COL_TEMPLATE] = "메로 반마리 {수량}개"
    preview = selection_preview(edited, ids, settings)
    assert preview.iloc[0]["수량 2"] == "메로 반마리 2개" and preview.iloc[0]["오류"] == ""
    outcome = register_rows(store, edited, USER, settings, ids)
    assert outcome.status == "saved" and len(outcome.appended.added) == 1
    dictionary, settings, _ = load_state(store)
    second = process_orders(ORDER, config, dictionary, settings)
    assert second.stats["rows_dictionary"] == 1 and second.stats["rows_needs_review"] == 1
    assert len(second.unmatched) == 1


def test_register_blocks_errors_and_needs_confirmation(config):
    store = make_store()
    ids = vendor_ids(config)
    sug = build_suggestions(process_orders(ORDER, config).unmatched, config, ids)
    edited = sug.copy()
    edited[COL_CHECK] = True
    edited[COL_TEMPLATE] = ""
    assert register_rows(store, edited, USER, SETTINGS, ids).status == "errors"
    assert load_dictionary_frame(store)[0].empty
    edited[COL_TEMPLATE] = "고정 문구"
    assert register_rows(store, edited, USER, SETTINGS, ids).status == "needs_confirm"
    assert register_rows(store, edited, USER, SETTINGS, ids, allow_warnings=True).status == "saved"
    assert register_rows(store, sug, USER, SETTINGS, ids).status == "nothing_selected"


def _work() -> pd.DataFrame:
    rows = [
        {"channel": "naver", "enabled": "1", "product_no": str(100 + i), "option_key": "", "product_name_ref": n,
         "vendor_id": "V1", "display_template": f"{n} {{수량}}개", "needs_review": "1"}
        for i, n in enumerate(["사과", "배", "포도"])
    ]
    frame = pd.DataFrame([{**dict.fromkeys(DICTIONARY_COLUMNS, ""), **r} for r in rows], columns=DICTIONARY_COLUMNS)
    return make_work(frame)


def test_edits_survive_filter_changes():
    work = _work()
    base = work.drop(columns=[DELETE])
    rids = filter_rids(work, "사과", "전체", False)
    assert rids == [0]
    view = build_view(work, rids, advanced=False)
    view.loc[0, V_TEMPLATE] = "사과 상자 {수량}"
    work = fold_edits(work, view)
    # change the filter: other rows appear and are edited; the earlier edit and hidden rows persist
    view2 = build_view(work, filter_rids(work, "", "전체", False), advanced=False)
    view2.loc[2, V_DELETE] = True
    work = fold_edits(work, view2)
    assert work.loc[0, "display_template"] == "사과 상자 {수량}"
    assert len(work) == 3 and work.loc[2, DELETE]
    assert changed_rids(base, work) == ([0], [2])
    table = change_table(base, work)
    assert list(table["구분"]) == ["수정", "삭제"] and "→" in table.iloc[0]["변경 내용"]
    saved = prepare_save(base, work, "me@example.com", now="NOW")
    assert list(saved["product_no"]) == ["100", "101"]
    assert saved.loc[0, "updated_by"] == "me@example.com" and saved.loc[1, "updated_by"] == ""


def test_fold_ignores_unchanged_formatting_and_hidden_rows():
    work = _work()
    work.loc[1, "enabled"] = "true"
    view = build_view(work, [1], advanced=True)
    assert fold_edits(work, view).equals(work)


def test_describe_issue_names_the_row():
    work = _work()
    work.loc[0, "vendor_id"] = "없음"
    kept = kept_rows(work)
    issue = validate_dictionary(kept, ["V1"], SETTINGS)[0]
    assert describe_issue(kept, issue).startswith("[사과 / 옵션 없음] 발주양식")
    assert RID not in kept.columns


def test_preview_bundle_dictionary_and_auto_rows(config):
    ids = vendor_ids(config)
    frame = pd.DataFrame([{**dict.fromkeys(DICTIONARY_COLUMNS, ""), "channel": "naver", "enabled": "1",
                           "product_no": "111", "option_key": "메로 선택: 반마리", "vendor_id": ids[4],
                           "display_template": "사전 메로 {수량}"}])
    dictionary = ItemDictionary(frame, SETTINGS)
    rows = pd.DataFrame({"상품번호": ["111", "999"], "상품명": ["메로 구이", "메로 찜"],
                         "옵션정보": ["메로 선택: 반마리", "메로 선택: 한마리"], "수량": [2, 1]})
    result = preview_bundle(rows, config, dictionary, SETTINGS)
    assert result.paths == ["사전", "자동 규칙"]
    assert len(result.vendors) == 1 and "사전 메로 2" in result.vendors[0].text
    assert list(result.debug) == [1] and result.debug[1]
    assert preview_bundle(rows.iloc[0:0], config, dictionary, SETTINGS).vendors == []


def test_bulk_set_only_touches_given_rows():
    from ui.dictionary_logic import (
        DELETE,
        V_DELETE,
        V_ENABLED,
        V_REVIEW,
        bulk_set,
        changed_rids,
        make_work,
    )

    base = make_work(pd.DataFrame([
        {"product_no": str(i), "option_key": "o", "vendor_id": "속초 발주양식", "display_template": "x",
         "enabled": "1", "needs_review": "1"} for i in range(4)
    ]))
    work = bulk_set(base, [1, 2], V_REVIEW, False)
    assert work["needs_review"].tolist() == ["1", "0", "0", "1"]
    work = bulk_set(work, [0, 3], V_ENABLED, False)
    assert work["enabled"].tolist() == ["0", "1", "1", "0"]
    work = bulk_set(work, [2], V_DELETE, True)
    assert work[DELETE].tolist() == [False, False, True, False]
    assert base["needs_review"].tolist() == ["1"] * 4  # input untouched
    modified, deleted = changed_rids(base.drop(columns=[DELETE]), work)
    assert deleted == [2] and sorted(modified) == [0, 1, 3]
