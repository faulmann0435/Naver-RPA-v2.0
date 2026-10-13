"""Synthetic order tables and xlsx files for the invoice tests (no real customer data)."""
from __future__ import annotations

from io import BytesIO

import openpyxl
import pandas as pd

ORDER_COLUMNS = ["상품주문번호", "주문번호", "배송방법", "택배사", "송장번호", "발송일", "수취인명", "우편번호",
                 "상품명", "수량", "수취인연락처1", "통합배송지", "배송비 묶음번호", "주문상태", "수취인연락처2"]
HANJIN_COLUMNS = ["순번", "도서", "운송장번호", "받으시는 분", "받으시는 분 전화", "받는분우편번호",
                  "받는분총주소", "품목명"]
ADDRESS = "서울특별시 가나구 다라로 1 (마바동) 101호"


def order_row(pid: str, name: str = "홍길동", **over) -> dict:
    row = {"상품주문번호": pid, "주문번호": "2026" + pid, "배송방법": "", "택배사": "", "송장번호": "",
           "발송일": "", "수취인명": name, "우편번호": "08239", "상품명": f"상품{pid}", "수량": 1,
           "수취인연락처1": "010-0000-0001", "통합배송지": ADDRESS, "배송비 묶음번호": f"B{pid}",
           "주문상태": "발송대기", "수취인연락처2": ""}
    return {**row, **over}


def orders(*rows: dict) -> pd.DataFrame:
    return pd.DataFrame(list(rows), columns=ORDER_COLUMNS)


def hanjin_row(seq: int, tracking: str, name: str = "홍길동", **over) -> dict:
    row = {"순번": str(seq), "도서": None, "운송장번호": tracking, "받으시는 분": name,
           "받으시는 분 전화": "01000000001", "받는분우편번호": "08239",
           "받는분총주소": ADDRESS + " [마바동, 12-3]", "품목명": "냉동식품"}
    return {**row, **over}


def _workbook_bytes(sheet_title: str, lines: list[list]) -> bytes:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = sheet_title
    for line in lines:
        sheet.append(line)
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def hanjin_xlsx(*rows: dict, columns: list[str] | None = None) -> bytes:
    columns = columns or HANJIN_COLUMNS
    return _workbook_bytes("sheet1", [columns] + [[r.get(c) for c in columns] for r in rows])


def order_export_xlsx(*rows: dict) -> bytes:
    """Like Naver's export: an instruction line first, then the header."""
    lines = [["안내: 이 줄은 설명입니다."] + [None] * (len(ORDER_COLUMNS) - 1), ORDER_COLUMNS]
    lines += [[r.get(c) for c in ORDER_COLUMNS] for r in rows]
    return _workbook_bytes("발주발송관리", lines)


def api_item(pid: str, name: str = "홍길동", package: str | None = None, status: str = "PAYED") -> dict:
    """A Naver API order detail whose recipient matches hanjin_row()'s defaults."""
    return {"order": {"orderId": "2026" + pid, "ordererName": "구매자", "ordererTel": "010-1111-2222"},
            "productOrder": {
                "productOrderId": pid, "productOrderStatus": status, "productName": f"상품{pid}",
                "productOption": "옵션", "quantity": 1, "packageNumber": package or f"9{pid}",
                "productId": "5", "shippingAddress": {
                    "name": name, "zipCode": "08239", "tel1": "010-0000-0001",
                    "baseAddress": "서울특별시 가나구 다라로 1 (마바동)", "detailedAddress": "101호"}}}
