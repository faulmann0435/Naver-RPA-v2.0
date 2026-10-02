"""⚙️ 고급 설정: product routing grid, option-rule grid, rule test and settings (tabs)."""
from __future__ import annotations

from typing import Literal

import pandas as pd
import requests
import streamlit as st

from core.engine import IMPLEMENTED_ACTIONS
from store.base import ConflictError, StoreError
from store.dictionary_repo import now_kst_iso
from ui import context, rules_state
from ui.rules_help import ACTION_HELP, ACTION_HELP_COLUMNS, ACTION_HELP_NOTE
from ui.rules_logic import (
    ALL_FILTER,
    OPTIONS_SPEC,
    RID,
    ROUTE_SPEC,
    V_ACTION,
    V_DELETE,
    V_DESC,
    V_ENABLED,
    V_KEYWORD,
    V_ORDER,
    V_PARAM,
    V_PRIORITY,
    V_TARGET,
    V_VENDOR,
    TableState,
    build_view,
    filter_option_rids,
    fold_grid,
    judge,
    sort_rids,
    summary_message,
    visible_rids,
    with_work,
)
from ui.rules_settings_tab import render_settings_tab
from ui.rules_tools import render_impact, render_rule_test


def _options_with_extras(base: list[str], values: pd.Series) -> list[str]:
    extras = [v for v in dict.fromkeys(values.astype(str)) if v and v not in base]
    return [*base, *extras]


def _grid(state: TableState, rids: list[int], num_rows: Literal["fixed", "dynamic"], columns: dict, key_extra: object) -> TableState:
    """Open the editor on a frozen copy of the rows and fold its result onto the matching working copy.

    The copy stays frozen while the filters and the working copy are only changed by this grid; when
    something else changed the working copy (or the filters changed) the editor gets a fresh key.
    """
    spec = state.spec
    version = st.session_state[rules_state.S_VERSION]
    frozen_key = rules_state.FROZEN_PREFIX + spec.kind
    frozen = st.session_state.get(frozen_key)
    signature = (version, num_rows, key_extra)
    if frozen is None or frozen["sig"] != signature or not frozen["result"].equals(state.work):
        epoch = frozen["epoch"] + 1 if frozen else 0
        frozen = {"sig": signature, "epoch": epoch, "work0": state.work, "view0": build_view(state.work, rids, spec)}
    edited = st.data_editor(
        frozen["view0"], key=f"rules_w_editor_{spec.kind}_{version}_{frozen['epoch']}", num_rows=num_rows,
        hide_index=True, column_config=columns, width="stretch",
    )
    result = fold_grid(frozen["work0"], frozen["view0"], edited, spec, allow_remove=num_rows == "dynamic")
    st.session_state[frozen_key] = {**frozen, "result": result}
    return with_work(state, result)


def _save(state: TableState, new_text: str, message: str) -> None:
    store = context.get_store()
    if store is None:
        return
    try:
        store.write_text(state.spec.file, new_text, state.sha, message, context.current_author())
    except ConflictError:
        st.session_state[rules_state.CONFLICT_PREFIX + state.spec.kind] = new_text
        return
    except (StoreError, requests.RequestException) as e:
        st.error(f"저장하지 못했습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    context.clear_data_caches()
    rules_state.load_all(store)
    st.session_state[rules_state.S_FLASH] = f"저장했습니다. {message}"
    st.rerun()


def _conflict_box(state: TableState) -> None:
    text = st.session_state.get(rules_state.CONFLICT_PREFIX + state.spec.kind)
    if text is None:
        return
    st.error("다른 사람이 먼저 저장했습니다. '새로 불러오기' 후 다시 수정하세요.")
    st.download_button(
        "내 수정본 받기", data=text.encode("utf-8"), file_name=f"{state.spec.file}_내수정본.csv",
        mime="text/csv", key=f"rules_w_mine_{state.spec.kind}",
    )


def _save_panel(state: TableState, other: TableState, ids: list[str], layout: pd.DataFrame) -> None:
    spec = state.spec
    user = context.current_user()
    judged = judge(state, state.work, ids, other.text, layout, user, now_kst_iso())
    st.caption(f"수정 {len(judged.modified)} · 추가 {len(judged.added)} · 삭제 {len(judged.deleted)}")
    clicked = st.button("저장", type="primary", disabled=not judged.has_changes, key=f"rules_w_save_{spec.kind}")
    for line in judged.errors:
        st.error(line)
    if judged.warnings:
        with st.expander(f"⚠ 고친 규칙의 경고 {len(judged.warnings)}건 (저장 전 확인)"):
            for line in judged.warnings:
                st.warning(line)
    if judged.reference:
        with st.expander(f"참고: 이번에 고치지 않은 규칙의 문제 {len(judged.reference)}건 (저장은 막지 않음)"):
            for line in judged.reference:
                st.caption(line)
    ask = bool(judged.warnings) and not judged.errors
    confirmed = st.checkbox("경고를 확인했습니다", key=f"rules_w_confirm_{spec.kind}") if ask else False
    if not clicked:
        return
    if judged.errors:
        st.error("오류가 있어 저장하지 않았습니다.")
    elif judged.warnings and not confirmed:
        st.warning("경고를 확인한 뒤 '경고를 확인했습니다'에 체크하고 다시 저장하세요.")
    else:
        _save(state, judged.new_text, summary_message(spec, judged))


def _route_tab(route: TableState, options: TableState, ids: list[str], layout: pd.DataFrame) -> TableState:
    st.caption("위에서부터 우선순위 숫자가 작은 줄을 먼저 검사합니다. 키워드: 쉼표로 여러 개, DEFAULT는 기본값")
    rids = sort_rids(route.work, visible_rids(route.work, ROUTE_SPEC), ROUTE_SPEC)
    vendors = _options_with_extras(ids, route.work["양식명칭"])
    column_config = {
        RID: None,
        V_ENABLED: st.column_config.CheckboxColumn("사용", default=True),
        V_PRIORITY: st.column_config.NumberColumn("우선순위", step=1, format="%d"),
        V_KEYWORD: st.column_config.TextColumn("키워드", help="쉼표로 여러 개, DEFAULT는 기본값"),
        V_VENDOR: st.column_config.SelectboxColumn("양식명칭", options=vendors),
    }
    route = _grid(route, rids, "dynamic", column_config, ())
    _save_panel(route, options, ids, layout)
    _conflict_box(route)
    return route


def _filters(ids: list[str]) -> tuple[str, str, str]:
    cols = st.columns([2, 2, 3])
    vendor = cols[0].selectbox("양식명칭", [ALL_FILTER, "ALL", *ids], key="rules_w_f_vendor")
    action = cols[1].selectbox("ActionType", [ALL_FILTER, *sorted(IMPLEMENTED_ACTIONS)], key="rules_w_f_action")
    query = cols[2].text_input("검색 (적용대상/Parameter/설명)", key="rules_w_f_query")
    return vendor, action, query


def _action_help() -> None:
    with st.expander("ActionType 설명"):
        st.dataframe(pd.DataFrame(ACTION_HELP, columns=ACTION_HELP_COLUMNS), hide_index=True, width="stretch")
        st.caption(ACTION_HELP_NOTE)


def _options_tab(options: TableState, route: TableState, ids: list[str], layout: pd.DataFrame) -> TableState:
    _action_help()
    vendor, action, query = _filters(ids)
    filtered = (vendor, action, query.strip()) != (ALL_FILTER, ALL_FILTER, "")
    rids = filter_option_rids(options.work, vendor, action, query)
    st.caption(
        f"표시 중 {len(rids)}개 / 전체 {len(visible_rids(options.work, OPTIONS_SPEC))}개 · "
        + ("필터를 켜면 줄 추가는 할 수 없습니다. " if filtered else "")
        + "순서 숫자를 바꾸면 저장할 때 1부터 다시 매겨집니다. 삭제는 '삭제' 칸에 체크하세요."
    )
    work = options.work
    column_config = {
        RID: None,
        V_ENABLED: st.column_config.CheckboxColumn("사용", default=True),
        V_ORDER: st.column_config.NumberColumn("순서", step=1, format="%d"),
        V_VENDOR: st.column_config.SelectboxColumn("양식명칭", options=_options_with_extras(["ALL", *ids], work["양식명칭"])),
        V_TARGET: st.column_config.TextColumn("적용대상", help="ALL 또는 문구 하나"),
        V_ACTION: st.column_config.SelectboxColumn(
            "ActionType", options=_options_with_extras(sorted(IMPLEMENTED_ACTIONS), work["ActionType (명령)"].str.upper())
        ),
        V_PARAM: st.column_config.TextColumn("Parameter"),
        V_DESC: st.column_config.TextColumn("설명"),
        V_DELETE: st.column_config.CheckboxColumn("삭제", default=False),
    }
    options = _grid(options, rids, "fixed" if filtered else "dynamic", column_config, (vendor, action, query))
    _save_panel(options, route, ids, layout)
    _conflict_box(options)
    return options


def render() -> None:
    st.subheader("⚙️ 고급 설정")
    st.warning("일반 작업은 '품목 관리'에서 하세요. 이 규칙은 품목 사전에 없는 품목에만 적용됩니다.")
    store = context.get_store()
    if store is None:
        st.info("데이터 저장소가 설정되지 않아 규칙을 편집할 수 없습니다. 설정(secrets)을 확인하세요.")
        return
    config = context.load_config_or_error()
    if config is None:
        return
    if st.button("새로 불러오기", key="rules_w_reload") or rules_state.S_ROUTE not in st.session_state:
        try:
            rules_state.load_all(store)
        except (StoreError, requests.RequestException, ValueError) as e:
            st.error(f"규칙을 불러올 수 없습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
            return
    flash = st.session_state.pop(rules_state.S_FLASH, None)
    if flash:
        st.success(flash)
    ids = context.vendor_ids(config)
    layout = context.load_raw_layout()
    route, options = rules_state.tables()
    tab_route, tab_options, tab_test, tab_settings = st.tabs(["상품분류", "옵션규칙", "규칙 시험", "설정"])
    with tab_route:
        route = _route_tab(route, options, ids, layout)
    with tab_options:
        options = _options_tab(options, route, ids, layout)
    st.session_state[rules_state.S_ROUTE], st.session_state[rules_state.S_OPTIONS] = route, options
    with tab_route:
        render_impact("route", config, route, options, layout)
    with tab_options:
        render_impact("options", config, route, options, layout)
    with tab_test:
        render_rule_test(config, route, options, layout)
    with tab_settings:
        render_settings_tab(store)
