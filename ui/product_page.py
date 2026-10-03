"""Product-by-product dictionary editing (tab 1 of 📖 품목 관리)."""
from __future__ import annotations

from collections.abc import Callable

import pandas as pd
import requests
import streamlit as st

from store.base import ConflictError, StoreError
from store.dictionary_repo import save_product
from ui import context
from ui.dictionary_logic import ALL_VENDORS
from ui.option_form import (
    clear_pending,
    init_state,
    option_form,
    state_key,
    vendor_select,
)
from ui.option_logic import FormValues, values_from_row
from ui.product_logic import (
    COL_CREATED,
    COL_DIRTY,
    COL_NAME,
    COL_OPTIONS,
    COL_REVIEW,
    DELETE,
    EditValue,
    ProductSummary,
    apply_product_edits,
    change_counts,
    common_vendor,
    dirty_products,
    group_key,
    option_prefix,
    product_issues,
    product_message,
    product_name,
    product_option_rows,
    product_summaries,
    summary_table,
)

S_SELECTED, S_DELETED, S_APPLIED = "pm_selected", "pm_deleted", "pm_applied_selection"
K_QUERY, K_VENDOR, K_REVIEW = "pm_query", "pm_vendor_filter", "pm_review_only"
TABLE_HEIGHT = 420
TABLE_COLUMNS = {
    COL_REVIEW: st.column_config.Column(width="small"),
    COL_OPTIONS: st.column_config.Column(width="small"),
    COL_DIRTY: st.column_config.Column(width="small"),
    COL_NAME: st.column_config.Column(width="large"),
    COL_CREATED: st.column_config.Column(width="small"),
}


def _table_key(summaries: list[ProductSummary]) -> str:
    return f"pm_table_{hash(tuple(s.group for s in summaries))}"


def _apply_table_selection(table_key: str, summaries: list[ProductSummary]) -> None:
    """A click on the list sets pm_selected (read before the right side is drawn)."""
    try:
        rows = tuple(st.session_state[table_key]["selection"]["rows"])
    except (KeyError, TypeError):
        return
    if rows and st.session_state.get(S_APPLIED) != (table_key, rows) and rows[0] < len(summaries):
        st.session_state[S_SELECTED] = summaries[rows[0]].group
    st.session_state[S_APPLIED] = (table_key, rows)


def _product_list(base: pd.DataFrame, ids: list[str], summaries: list[ProductSummary], table_key: str) -> None:
    st.text_input("검색 (상품명/옵션)", key=K_QUERY)
    counts = base["vendor_id"].fillna("").astype(str).str.strip().value_counts().to_dict()
    st.selectbox(
        "발주양식", [ALL_VENDORS, *ids], key=K_VENDOR,
        format_func=lambda v: v if v == ALL_VENDORS else f"{v} ({counts.get(v, 0)})",
    )
    st.checkbox("확인 안 한 것만", key=K_REVIEW)
    dirty = dirty_products(base, st.session_state.get("pending", {}), st.session_state.get(S_DELETED, frozenset()))
    st.caption(f"상품 {len(summaries)}개")
    vendor = st.session_state.get(K_VENDOR, ALL_VENDORS)
    if not summaries and vendor != ALL_VENDORS and not counts.get(vendor, 0):
        st.info(
            f"'{vendor}'로 등록된 품목이 아직 없습니다. 이 발주양식으로 가는 주문이 들어오면 "
            "주문처리 화면의 '사전 미등록' 상자에서 등록할 수 있고, 등록하면 여기에 나타납니다."
        )
    st.dataframe(
        summary_table(summaries, dirty), key=table_key, on_select="rerun", selection_mode="single-row",
        hide_index=True, width="stretch", height=TABLE_HEIGHT, column_config=TABLE_COLUMNS,
    )


# ---------------------------------------------------------------- right side

def _sync_product_vendor(group: str, prefixes: list[str], ids: list[str]) -> None:
    """The product-level 발주양식: a change moves every option that still follows the old value."""
    state = st.session_state
    pv_key, prev_key = vendor_keys(group)
    vendors = [state[state_key(p, "vendor")] for p in prefixes]
    if pv_key not in state:
        state[pv_key] = state[prev_key] = common_vendor([{"vendor_id": v} for v in vendors])
    current = state[pv_key]
    options = ids if current in ids else [*ids, current]
    chosen = st.selectbox("상품 발주양식", options, key=pv_key)
    previous = state.get(prev_key, chosen)
    if chosen != previous:
        for prefix in prefixes:
            if state[state_key(prefix, "vendor")] == previous:
                state[state_key(prefix, "vendor")] = chosen
    state[prev_key] = chosen
    if len({state[state_key(p, "vendor")] for p in prefixes}) > 1:
        st.caption("옵션마다 다름")


def vendor_keys(group: str) -> tuple[str, str]:
    """Session keys of the product-level vendor select and its previous value."""
    return f"pm_pv_{group_key(group)}", f"pm_pv_prev_{group_key(group)}"


def _option_label(row: dict[str, str]) -> str:
    return row["option_raw_ref"] or row["option_key"] or "옵션 없음"


def _reset_product(group: str, prefixes: list[str]) -> None:
    for prefix in prefixes:
        clear_pending(prefix)
    st.session_state[S_DELETED] = frozenset(st.session_state.get(S_DELETED, frozenset())) - set(prefixes)
    for key in vendor_keys(group):
        st.session_state.pop(key, None)


def _option_blocks(rows: list[dict[str, str]], ids: list[str]) -> tuple[dict[str, FormValues], set[str]]:
    values: dict[str, FormValues] = {}
    deleted: set[str] = set()
    for row in rows:
        prefix = option_prefix(row["product_no"], row["option_key"])
        with st.container(border=True):
            st.markdown(f"**{_option_label(row)}**")
            values[prefix] = option_form(prefix, values_from_row(row), ids, show_vendor=False, show_flags=True)
            with st.expander("삭제"):
                if st.checkbox("이 옵션을 사전에서 삭제", key=state_key(prefix, "delete")):
                    deleted.add(prefix)
    return values, deleted


def _save(
    base: pd.DataFrame, sha: str | None, group: str, name: str, edits: dict[tuple[str, str], EditValue],
    ids: list[str], reload: Callable[[], None], prefixes: list[str],
) -> None:
    new = apply_product_edits(base, group, edits, user=context.current_user())
    modified, deleted = change_counts(base, new, group)
    if modified + deleted == 0:
        st.info("바뀐 내용이 없습니다.")
        return
    state = context.load_dictionary_state_or_error()
    if state is None:
        return
    errors, _warnings = product_issues(new, group, ids, state[1])
    if errors:
        for message in errors:
            st.error(message)
        st.error("오류가 있어 저장하지 않았습니다.")
        return
    message = product_message(name, modified, deleted)
    try:
        save_product(context.get_store(), base, sha, new, group, context.current_author(), message)  # type: ignore[arg-type]
    except ConflictError:
        st.error("다른 사람이 이 상품을 먼저 수정했습니다. [새로 불러오기]를 눌러 최신 내용을 확인한 뒤 다시 고치세요.")
        return
    except (StoreError, requests.RequestException) as e:
        st.error(f"저장하지 못했습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    context.clear_data_caches()
    reload()
    _reset_product(group, prefixes)
    st.session_state["dict_flash"] = f"저장했습니다. {message}"
    st.rerun()


def _detail(base: pd.DataFrame, sha: str | None, ids: list[str], group: str, reload: Callable[[], None]) -> None:
    rows = product_option_rows(base, group)
    if not rows:
        st.info("이 상품은 사전에 없습니다. 목록에서 다시 고르세요.")
        return
    name = product_name(rows)
    prefixes = [option_prefix(r["product_no"], r["option_key"]) for r in rows]
    st.markdown(f"### {name}")
    deleted_before = st.session_state.get(S_DELETED, frozenset())
    for row, prefix in zip(rows, prefixes, strict=True):
        init_state(prefix, values_from_row(row))
        st.session_state.setdefault(state_key(prefix, "delete"), prefix in deleted_before)
    _sync_product_vendor(group, prefixes, ids)
    with st.expander("옵션별 발주양식 다르게"):
        for row, prefix in zip(rows, prefixes, strict=True):
            vendor_select(prefix, values_from_row(row), ids, label=_option_label(row))
    values, deleted = _option_blocks(rows, ids)
    st.session_state[S_DELETED] = (frozenset(deleted_before) - set(prefixes)) | deleted
    cols = st.columns([1, 1, 3])
    save = cols[0].button("이 상품 저장", type="primary", key="pm_save")
    if cols[1].button("변경 취소", key="pm_cancel"):
        _reset_product(group, prefixes)
        st.rerun()
    if save:
        edits: dict[tuple[str, str], EditValue] = {
            (row["product_no"], row["option_key"]): DELETE if prefix in deleted else values[prefix]
            for row, prefix in zip(rows, prefixes, strict=True)
        }
        _save(base, sha, group, name, edits, ids, reload, prefixes)


def render_tab(base: pd.DataFrame, sha: str | None, ids: list[str], reload: Callable[[], None]) -> None:
    summaries = product_summaries(
        base, st.session_state.get(K_QUERY, ""), st.session_state.get(K_VENDOR, ALL_VENDORS),
        bool(st.session_state.get(K_REVIEW, False)),
    )
    table_key = _table_key(summaries)
    _apply_table_selection(table_key, summaries)
    left, right = st.columns([1, 2])
    # The right side is drawn first so the list (left) shows the "저장 안 됨" mark of the current run.
    with right:
        selected = st.session_state.get(S_SELECTED)
        if selected:
            _detail(base, sha, ids, str(selected), reload)
        else:
            st.info("왼쪽 목록에서 상품을 고르세요.")
    with left:
        _product_list(base, ids, summaries, table_key)
