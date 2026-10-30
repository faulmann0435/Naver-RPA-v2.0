"""🚚 송장 입력: Naver order export + Hanjin tracking file -> the upload xlsx for 엑셀 일괄발송."""
from __future__ import annotations

from datetime import datetime

import requests
import streamlit as st

from core.invoice_match import KST, Assignment
from core.loader import HAS_MSOFFCRYPTO
from store import dispatch_queue
from store.base import DataStore, StoreError
from ui import context
from ui import invoice_logic as logic

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
REQUEST_ID_KEY = "invoice_request_id"
SUBMITTED_KEY = "invoice_submitted_keys"
REFRESH_SECONDS = 10
STORE_ERRORS = (StoreError, requests.RequestException, ValueError, KeyError)


def _tables(result) -> None:
    sections = (
        ("일치", logic.matched_table(result), True),
        ("확인 필요", logic.check_table(result), True),
        ("주문을 못 찾은 송장", logic.loose_invoice_table(result), False),
        ("송장번호가 없는 줄", logic.no_tracking_table(result), True),
        ("송장 없는 발송대기 주문", logic.unshipped_table(result), False),
    )
    for title, table, expanded in sections:
        with st.expander(f"{title} ({len(table)})", expanded=expanded and not table.empty):
            if table.empty:
                st.caption("없습니다.")
            else:
                st.dataframe(table, hide_index=True, width="stretch")


def _download(outcome: logic.InvoiceOutcome) -> tuple[Assignment, ...]:
    """The download button; returns the selected shipments (empty when there is nothing to upload)."""
    extra = len(outcome.result.includable_checks())
    include = st.checkbox(f"확인 필요 중 포함 가능한 건도 넣기 ({extra}건)", value=False,
                          disabled=not extra, key="invoice_include_checks")
    name, data = logic.build_download(outcome, include)
    selected = outcome.result.includable() if include else outcome.result.matched()
    if not selected:
        st.info("올릴 송장이 없습니다.")
        return ()
    st.download_button(f"📥 {name} ({len(selected)}건)", data=data, file_name=name, mime=EXCEL_MIME)
    st.caption(logic.NOTE_INSTRUCTION)
    return selected


@st.fragment(run_every=REFRESH_SECONDS)
def _request_status(store: DataStore) -> None:
    """Shows the remembered request and refreshes itself every few seconds."""
    request_id = st.session_state.get(REQUEST_ID_KEY)
    if not request_id:
        return
    try:
        request = dispatch_queue.get(store, request_id)
    except STORE_ERRORS as e:
        st.caption(f"처리 상태를 불러오지 못했습니다. 곧 다시 시도합니다. ({type(e).__name__})")
        return
    if request is None:
        st.caption("요청을 찾을 수 없습니다.")
        return
    show = {"info": st.info, "warning": st.warning, "success": st.success, "error": st.error}
    for level, text in logic.status_messages(request, datetime.now(KST)):
        show[level](text)


def _register(selected: tuple[Assignment, ...]) -> None:
    """Section 'register on Naver now': queues the selection for the owner's PC."""
    st.markdown("#### 🚀 네이버에 바로 등록 (사장님 PC가 처리)")
    store = context.get_store()
    if store is None:
        st.info(logic.REGISTER_NEEDS_STORE)
        return
    items = logic.queue_items(selected) if selected else []
    key = logic.selection_key(items) if items else ""
    submitted: set[str] = st.session_state.setdefault(SUBMITTED_KEY, set())
    if items and st.button(f"🚀 네이버에 송장 등록 ({len(selected)}건)", disabled=key in submitted,
                           key="invoice_register_button"):
        try:
            st.session_state[REQUEST_ID_KEY] = dispatch_queue.submit(
                store, items, context.current_user(), datetime.now(KST))
        except STORE_ERRORS as e:
            st.error(f"요청을 보내지 못했습니다: {e}")
        else:
            submitted.add(key)
            st.info(logic.REGISTER_SENT)
    _request_status(store)


def render() -> None:
    st.subheader("🚚 송장 입력")
    st.write("주문서 엑셀과 송장 파일을 올리면 받는 사람이 같은 주문끼리 송장번호를 맞춰 줍니다.")
    if not HAS_MSOFFCRYPTO:
        st.warning("비밀번호 보호 엑셀: `pip install msoffcrypto-tool`")
    order_file = st.file_uploader("주문서 엑셀 (스마트스토어 발주발송관리)", type=["xlsx"], key="invoice_orders")
    invoice_files = st.file_uploader("송장 파일 (한진 송장파일, 한진 예약양식, 공급처 송장 목록, '한진택배' 칸을 채운 발주서 · 여러 개 가능)",
                                     type=["xlsx", "xls"], accept_multiple_files=True, key="invoice_hanjin")
    password = st.text_input("주문서 엑셀 비밀번호 (없으면 비움)", type="password", key="invoice_pw")
    if order_file is None or not invoice_files:
        st.info("주문서와 송장 파일을 모두 올려 주세요.")
        return
    try:
        outcome = logic.match_files(order_file.getvalue(), [f.getvalue() for f in invoice_files], password)
    except ValueError as e:
        st.error(str(e))
        return
    for column, (label, value) in zip(st.columns(4), logic.metrics(outcome.result).items(), strict=True):
        column.metric(label, value)
    blank = len(outcome.result.invoices_without_tracking)
    if blank:
        st.warning(f"송장번호가 없는 줄이 {blank}개 있습니다. 보낸 곳에 송장을 빠뜨리지 않았는지 확인하세요.")
    _tables(outcome.result)
    _register(_download(outcome))
