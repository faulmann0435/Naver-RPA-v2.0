"""📖 품목 관리: product-by-product editing (tab 1) and the whole-dictionary grid editor (tab 2)."""
from __future__ import annotations

import pandas as pd
import requests
import streamlit as st

from store.base import ConflictError, StoreError
from store.csv_codec import to_csv_text
from store.dictionary_repo import save_dictionary
from store.validators import Issue, validate_dictionary
from ui import context, product_page
from ui.dictionary_logic import (
    ALL_VENDORS,
    BULK_FIELDS,
    DELETE,
    DISABLED_VIEW,
    RID,
    V_APPEND,
    V_CREATED,
    V_DELETE,
    V_ENABLED,
    V_GROUP,
    V_QTY1,
    V_REVIEW,
    V_TEMPLATE,
    V_VENDOR,
    V_WEIGHT,
    build_view,
    bulk_set,
    change_table,
    changed_rids,
    describe_issue,
    filter_rids,
    fold_edits,
    kept_rows,
    make_work,
    prepare_save,
    to_excel_bytes,
)
from ui.excel_import_page import render_import

S_BASE, S_WORK, S_VERSION, S_CONFLICT, S_XLSX = "dict_base", "dict_work", "dict_version", "dict_conflict", "dict_xlsx"


def load_base() -> None:
    """(Re)load dictionary.csv into the base and the working copy."""
    frame, sha = context.load_dictionary_frame_fresh()
    base = make_work(frame).drop(columns=[DELETE])
    st.session_state[S_BASE] = (base, sha)
    st.session_state[S_WORK] = make_work(frame)
    st.session_state[S_VERSION] = st.session_state.get(S_VERSION, 0) + 1
    st.session_state.pop(S_CONFLICT, None)


def _editor_config(ids: list[str]) -> dict:
    return {
        RID: None,
        V_ENABLED: st.column_config.CheckboxColumn(V_ENABLED),
        V_REVIEW: st.column_config.CheckboxColumn(V_REVIEW),
        V_CREATED: st.column_config.TextColumn(V_CREATED, width="small"),
        V_VENDOR: st.column_config.SelectboxColumn(V_VENDOR, options=ids, required=True),
        V_TEMPLATE: st.column_config.TextColumn(V_TEMPLATE),
        V_QTY1: st.column_config.TextColumn(V_QTY1),
        V_GROUP: st.column_config.TextColumn(V_GROUP),
        V_WEIGHT: st.column_config.NumberColumn(V_WEIGHT, min_value=0.0, format="%.3f"),
        V_APPEND: st.column_config.CheckboxColumn(V_APPEND),
        V_DELETE: st.column_config.CheckboxColumn(V_DELETE, default=False),
    }


def _filters(ids: list[str]) -> tuple[str, str, bool, bool]:
    cols = st.columns([3, 2, 1, 1])
    query = cols[0].text_input("검색 (상품명/옵션/표기)", key="dict_query")
    vendor = cols[1].selectbox("발주양식", [ALL_VENDORS, *ids], key="dict_vendor")
    review = cols[2].checkbox("검토 필요만", key="dict_review")
    advanced = cols[3].toggle("고급 칸 보기", key="dict_advanced")
    return query, vendor, review, advanced


def _bulk_controls(work: pd.DataFrame, rids: list[int]) -> None:
    """Check / uncheck one checkbox column for every row currently shown (after filters)."""
    cols = st.columns([2, 1, 1, 4], vertical_alignment="bottom")
    label = cols[0].selectbox("표시된 항목 일괄 변경", list(BULK_FIELDS), key="dict_bulk_target")
    value = True if cols[1].button("모두 체크", key="dict_bulk_on") else False if cols[2].button(
        "모두 해제", key="dict_bulk_off") else None
    cols[3].caption(f"지금 표에 보이는 {len(rids)}개 항목에만 적용됩니다. [저장]을 눌러야 반영됩니다.")
    if value is None or not rids:
        return
    st.session_state[S_WORK] = bulk_set(work, rids, label, value)
    st.session_state[S_VERSION] += 1  # new editor key so the table shows the bulk change
    st.rerun()


def _editor(work: pd.DataFrame, ids: list[str]) -> pd.DataFrame:
    query, vendor, review, advanced = _filters(ids)
    rids = filter_rids(work, query, vendor, review)
    st.caption(f"표시 중 {len(rids)}개 / 전체 {len(work)}개")
    _bulk_controls(work, rids)
    # The key changes with the filters and the row set, so stale edit positions are never replayed.
    key = f"dict_editor_{st.session_state[S_VERSION]}_{hash((query, vendor, review, advanced, tuple(rids)))}"
    edited = st.data_editor(
        build_view(work, rids, advanced), key=key, num_rows="fixed", hide_index=True,
        column_config=_editor_config(ids), disabled=DISABLED_VIEW, width="stretch",
    )
    return fold_edits(work, edited)


def _preview(base: pd.DataFrame, work: pd.DataFrame) -> None:
    modified, deleted = changed_rids(base, work)
    st.write(f"수정 {len(modified)}, 삭제 {len(deleted)}")
    st.dataframe(change_table(base, work), hide_index=True, width="stretch")


def _save(base: pd.DataFrame, sha: str | None, work: pd.DataFrame) -> None:
    modified, deleted = changed_rids(base, work)
    out = prepare_save(base, work, context.current_user())
    message = f"사전 수정: 수정 {len(modified)}, 삭제 {len(deleted)}"
    try:
        save_dictionary(context.get_store(), out, sha, context.current_author(), message)  # type: ignore[arg-type]
    except ConflictError:
        st.session_state[S_CONFLICT] = to_csv_text(kept_rows(work))
        return
    except (StoreError, requests.RequestException) as e:
        st.error(f"저장하지 못했습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    context.clear_data_caches()
    load_base()
    st.session_state["dict_flash"] = f"저장했습니다. {message}"
    st.rerun()


def _conflict_box() -> None:
    csv_text = st.session_state.get(S_CONFLICT)
    if csv_text is None:
        return
    st.error("다른 사람이 먼저 저장했습니다. 새로 불러온 뒤 다시 수정하세요.")
    st.download_button("내 수정본 받기", data=csv_text.encode("utf-8"), file_name="dictionary_내수정본.csv",
                       mime="text/csv", key="dict_mine")


def _save_controls(base: pd.DataFrame, sha: str | None, work: pd.DataFrame, ids: list[str]) -> None:
    state = context.load_dictionary_state_or_error()
    if state is None:
        return
    kept = kept_rows(work)
    modified, _deleted = changed_rids(base, work)
    has_changes = any(changed_rids(base, work))
    # Only the rows being changed are judged: their errors block the save and their warnings need a
    # confirmation. Problems in untouched rows (e.g. the draft) are listed for reference only.
    issues = validate_dictionary(kept, ids, state[1]) if has_changes else []
    touched = set(modified)

    def _is_touched(issue: Issue) -> bool:
        return issue.row is None or int(kept.index[issue.row]) in touched

    errors = [describe_issue(kept, i) for i in issues if i.level == "error" and _is_touched(i)]
    warnings = [describe_issue(kept, i) for i in issues if i.level == "warning" and _is_touched(i)]
    other_errors = [describe_issue(kept, i) for i in issues if i.level == "error" and not _is_touched(i)]
    cols = st.columns(2)
    if cols[0].button("변경사항 미리보기"):
        _preview(base, work)
    clicked = cols[1].button("저장", type="primary", disabled=not has_changes)
    for message in errors:
        st.error(message)
    if warnings:
        with st.expander(f"⚠ 고친 항목의 경고 {len(warnings)}건 (저장 전 확인)"):
            for message in warnings:
                st.warning(message)
    if other_errors:
        with st.expander(f"참고: 이번에 고치지 않은 항목의 오류 {len(other_errors)}건 (저장은 막지 않음)"):
            for message in other_errors:
                st.caption(message)
    confirmed = st.checkbox("경고를 확인했습니다", key="dict_confirm") if warnings and not errors else False
    if clicked and errors:
        st.error("오류가 있어 저장하지 않았습니다.")
    elif clicked and warnings and not confirmed:
        st.warning("경고를 확인한 뒤 '경고를 확인했습니다'에 체크하고 다시 저장하세요.")
    elif clicked:
        _save(base, sha, work)


def _excel_export(base: pd.DataFrame, sha: str | None) -> None:
    cached = st.session_state.get(S_XLSX)
    if cached is None or cached[0] != sha:
        cached = (sha, to_excel_bytes(base))
        st.session_state[S_XLSX] = cached
    st.download_button(
        "엑셀로 내보내기", data=cached[1], file_name="품목사전.xlsx", key="dict_xlsx_dl",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _grid_tab(base: pd.DataFrame, sha: str | None, ids: list[str]) -> None:
    """The whole dictionary as one grid: filters, bulk change, save and Excel export."""
    st.caption(f"사전 {len(base)}개 항목 · 버전 {(sha or '없음')[:7]} 기준으로 편집 중")
    work = _editor(st.session_state[S_WORK], ids)
    st.session_state[S_WORK] = work
    _save_controls(base, sha, work, ids)
    _conflict_box()
    _excel_export(base, sha)
    render_import(base, sha, ids, load_base)


def render() -> None:
    st.subheader("📖 품목 관리")
    st.caption("상품을 고르고, 옵션마다 발주서에 적을 문장을 쓰세요. 수량에 따라 늘어나는 숫자에 체크하면 됩니다.")
    if context.get_store() is None:
        st.info("데이터 저장소가 설정되지 않아 품목사전을 사용할 수 없습니다. 설정(secrets)을 확인하세요.")
        return
    config = context.load_config_or_error()
    if config is None:
        return
    refresh = st.button("새로 불러오기")
    if S_BASE not in st.session_state or refresh:
        try:
            load_base()
        except (StoreError, requests.RequestException) as e:
            st.error(f"품목사전을 불러올 수 없습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
            return
    base, sha = st.session_state[S_BASE]
    flash = st.session_state.pop("dict_flash", None)
    if flash:
        st.success(flash)
    ids = context.vendor_ids(config)
    product_tab, grid_tab = st.tabs(["상품별 편집", "표로 한꺼번에 보기"])
    with product_tab:
        product_page.render_tab(base, sha, ids, load_base)
    with grid_tab:
        _grid_tab(base, sha, ids)
