"""🚚 송장 입력: Naver order export + Hanjin tracking file -> the upload xlsx for 엑셀 일괄발송."""
from __future__ import annotations

import streamlit as st

from core.loader import HAS_MSOFFCRYPTO
from ui import invoice_logic as logic

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


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


def _download(outcome: logic.InvoiceOutcome) -> None:
    extra = len(outcome.result.includable_checks())
    include = st.checkbox(f"확인 필요 중 포함 가능한 건도 넣기 ({extra}건)", value=False,
                          disabled=not extra, key="invoice_include_checks")
    name, data = logic.build_download(outcome, include)
    count = len(outcome.result.includable() if include else outcome.result.matched())
    if not count:
        st.info("올릴 송장이 없습니다.")
        return
    st.download_button(f"📥 {name} ({count}건)", data=data, file_name=name, mime=EXCEL_MIME)
    st.caption(logic.NOTE_INSTRUCTION)


def render() -> None:
    st.subheader("🚚 송장 입력")
    st.write("주문서 엑셀과 송장 파일을 올리면 받는 사람이 같은 주문끼리 송장번호를 맞춰 줍니다.")
    if not HAS_MSOFFCRYPTO:
        st.warning("비밀번호 보호 엑셀: `pip install msoffcrypto-tool`")
    order_file = st.file_uploader("주문서 엑셀 (스마트스토어 발주발송관리)", type=["xlsx"], key="invoice_orders")
    invoice_files = st.file_uploader("송장 파일 (한진 송장파일, 또는 '한진택배' 칸에 송장번호를 적은 발주서 · 여러 개 가능)",
                                     type=["xlsx"], accept_multiple_files=True, key="invoice_hanjin")
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
    _download(outcome)
