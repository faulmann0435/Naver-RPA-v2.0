"""📜 송장 등록 기록: every 송장 registration, started from Telegram or from the site (dispatch/queue.json)."""
from __future__ import annotations

import requests
import streamlit as st

from store import dispatch_queue
from store.base import StoreError
from ui import context
from ui import dispatch_log_logic as logic

STORE_ERRORS = (StoreError, requests.RequestException, ValueError, KeyError)
FILTER_KEY = "dispatch_log_filter"
NEEDS_STORE = "기록을 보려면 데이터 저장소 설정이 필요합니다."


def _details(request: dict) -> None:
    with st.expander(logic.expander_title(request)):
        error = (request.get("result") or {}).get("error")
        if request.get("status") == dispatch_queue.STATUS_FAILED and error:
            st.error(f"등록하지 못했습니다: {error}")
        table = logic.detail_table(request)
        if table.empty:
            st.caption("송장 내용이 없습니다.")
        else:
            st.dataframe(table, hide_index=True, width="stretch")


def render() -> None:
    st.subheader("📜 송장 등록 기록")
    st.caption(f"텔레그램과 사이트에서 한 송장 등록을 모두 보여 줍니다. 최신순, 최근 {dispatch_queue.MAX_REQUESTS}건까지.")
    store = context.get_store()
    if store is None:
        st.info(NEEDS_STORE)
        return
    left, right = st.columns([4, 1])
    choice = left.radio("경로", logic.FILTER_CHOICES, horizontal=True, key=FILTER_KEY, label_visibility="collapsed")
    right.button("🔄 새로고침")  # any click reruns the page, which reads the log again
    try:
        everything = dispatch_queue.history(store)
    except STORE_ERRORS as e:
        st.error(f"기록을 불러오지 못했습니다. 잠시 뒤 다시 시도하세요. ({type(e).__name__})")
        return
    shown = logic.filter_by_source(everything, choice)
    if not shown:
        st.info("아직 기록이 없습니다.")
        return
    st.dataframe(logic.summary_table(shown), hide_index=True, width="stretch")
    for request in shown:
        _details(request)
