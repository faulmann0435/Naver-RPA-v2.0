"""송장 입력 page logic: Naver order export + Hanjin file -> matching tables and the upload xlsx.

Pure functions (no streamlit); the page in ui/invoice_page.py only shows what they return.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from io import BytesIO

import pandas as pd

from core.invoice_match import (
    KST,
    Assignment,
    MatchResult,
    build_upload_workbook,
    match_invoices,
    read_invoice_file,
    read_order_export,
)
from core.loader import _get_excel_bytes

UPLOAD_FILE_NAME = "송장업로드_{when:%Y%m%d_%H%M}.xlsx"
NOTE_INSTRUCTION = ("내려받은 파일을 스마트스토어센터 판매관리 > 발주(주문)확인/발송관리 > "
                    "엑셀 일괄발송에서 올리세요.")


@dataclass(frozen=True)
class InvoiceOutcome:
    orders: pd.DataFrame
    result: MatchResult


def match_files(order_bytes: bytes, invoice_bytes: bytes, password: str | None = None) -> InvoiceOutcome:
    """Decrypt and read both files, then match. Raises ValueError with a Korean message."""
    plain = _get_excel_bytes(BytesIO(order_bytes), password=password)
    orders = read_order_export(plain)
    return InvoiceOutcome(orders, match_invoices(orders, read_invoice_file(invoice_bytes)))


def metrics(result: MatchResult) -> dict[str, int]:
    return {
        "일치": len(result.matched()),
        "확인 필요": len(result.needs_check()),
        "주문 못 찾은 송장": len(result.invoices_without_order),
        "송장 없는 주문": len(result.orders_without_invoice),
    }


def _assignment_rows(assignments: tuple[Assignment, ...]) -> list[dict]:
    return [{"수취인": a.recipient, "상품": a.items, "송장번호": a.tracking,
             "상품주문번호": ", ".join(a.product_order_ids)} for a in assignments]


def matched_table(result: MatchResult) -> pd.DataFrame:
    return pd.DataFrame(_assignment_rows(result.matched()), columns=["수취인", "상품", "송장번호", "상품주문번호"])


def check_table(result: MatchResult) -> pd.DataFrame:
    rows = [{**row, "사유": a.reason, "포함 가능": "예" if a.includable else "아니오"}
            for a, row in zip(result.needs_check(), _assignment_rows(result.needs_check()), strict=True)]
    return pd.DataFrame(rows, columns=["수취인", "사유", "포함 가능", "송장번호", "상품", "상품주문번호"])


def loose_invoice_table(result: MatchResult) -> pd.DataFrame:
    rows = [{"순번": i.seq, "받는분": i.name, "송장번호": i.tracking, "품목": i.item_name}
            for i in result.invoices_without_order]
    return pd.DataFrame(rows, columns=["순번", "받는분", "송장번호", "품목"])


def unshipped_table(result: MatchResult) -> pd.DataFrame:
    rows = [{"수취인": o.recipient, "상품": o.items, "상품주문번호": ", ".join(o.product_order_ids)}
            for o in result.orders_without_invoice]
    return pd.DataFrame(rows, columns=["수취인", "상품", "상품주문번호"])


def build_download(outcome: InvoiceOutcome, include_checks: bool, when: datetime | None = None) -> tuple[str, bytes]:
    """(file name, xlsx bytes): matched rows, plus the includable checks when asked."""
    assignments = outcome.result.includable() if include_checks else outcome.result.matched()
    moment = when or datetime.now(KST)
    return UPLOAD_FILE_NAME.format(when=moment), build_upload_workbook(outcome.orders, assignments)
