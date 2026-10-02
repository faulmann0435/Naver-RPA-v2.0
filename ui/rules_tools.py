"""고급 설정 page parts: 변경 영향 미리보기 and the 규칙 시험 tab."""
from __future__ import annotations

import pandas as pd
import streamlit as st

from store.dictionary_repo import now_kst_iso
from ui import context
from ui.order_page import S_DF
from ui.rules_extra_logic import run_impact, run_rule_test
from ui.rules_logic import TableState, edited_config

CONFIG_ERRORS = (ValueError, KeyError, TypeError, AttributeError)
MAX_TABLE_ROWS = 50


def _edited(route: TableState, options: TableState, layout: pd.DataFrame) -> dict | None:
    try:
        return edited_config(route, options, layout, context.current_user(), now_kst_iso())
    except CONFIG_ERRORS as e:
        st.error(f"편집 중인 규칙으로는 처리할 수 없습니다. ({type(e).__name__}: {str(e)[:200]})")
        return None


def render_impact(kind: str, config: dict, route: TableState, options: TableState, layout: pd.DataFrame) -> None:
    """Process the order file of this session with the saved and the edited rules and compare."""
    st.markdown("**변경 영향 미리보기**")
    orders = st.session_state.get(S_DF)
    if orders is None:
        st.caption("먼저 '주문처리'에서 주문 파일을 처리해야 변경 영향을 볼 수 있습니다.")
        return
    if not st.button("변경 영향 미리보기", key=f"rules_w_impact_{kind}"):
        return
    state = context.load_dictionary_state_or_error()
    edited = _edited(route, options, layout)
    if state is None or edited is None:
        return
    try:
        result = run_impact(orders, config, edited, state[0], state[1])
    except CONFIG_ERRORS as e:
        st.error(f"미리보기에 실패했습니다: {type(e).__name__}: {str(e)[:200]}")
        return
    st.metric("바뀌는 발주서 줄", f"{result.changed_lines} / 전체 {result.total_lines}")
    if result.table.empty:
        st.success("저장했을 때 달라지는 발주서 줄이 없습니다.")
        return
    st.dataframe(result.table.head(MAX_TABLE_ROWS), hide_index=True, width="stretch")
    if len(result.table) > MAX_TABLE_ROWS:
        st.caption(f"{MAX_TABLE_ROWS}줄까지만 보여줍니다. (전체 {len(result.table)}줄)")


def render_rule_test(config: dict, route: TableState, options: TableState, layout: pd.DataFrame) -> None:
    """One item through routing and the auto rules (the dictionary is NOT used)."""
    st.caption("품목 사전은 쓰지 않고, 자동 규칙만으로 어떻게 나오는지 시험합니다.")
    cols = st.columns([3, 3, 1])
    name = cols[0].text_input("상품명", key="rules_w_t_name")
    option = cols[1].text_input("옵션정보", key="rules_w_t_option")
    qty = cols[2].number_input("수량", min_value=1, value=1, step=1, key="rules_w_t_qty")
    use_edit = st.checkbox("편집 중인 규칙으로 시험", key="rules_w_t_edit")
    if not name.strip() and not option.strip():
        st.info("상품명이나 옵션정보를 입력하세요.")
        return
    active = _edited(route, options, layout) if use_edit else config
    if active is None:
        return
    try:
        result = run_rule_test(name, option, int(qty), active)
    except CONFIG_ERRORS as e:
        st.error(f"시험에 실패했습니다: {type(e).__name__}: {str(e)[:200]}")
        return
    if result is None:
        st.info("이 입력으로는 발주서 줄이 만들어지지 않습니다.")
        return
    st.markdown(f"**분류된 발주양식:** {result.vendor}")
    st.markdown(f"**최종 문구:** `{result.text}`")
    applied_only = st.checkbox("적용된 규칙만 보기", value=True, key="rules_w_t_applied")
    steps = [s for s in result.steps if s.applied or not applied_only]
    table = pd.DataFrame(
        [{"번호": s.number, "ActionType": s.action, "Parameter": s.param, "양식": s.vendor_rule,
          "적용대상": s.target, "설명": s.description, "결과": s.note} for s in steps],
        columns=["번호", "ActionType", "Parameter", "양식", "적용대상", "설명", "결과"],
    )
    st.dataframe(table, hide_index=True, width="stretch")
