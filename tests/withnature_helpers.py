"""Synthetic 자연이랑 출고작업관리 files (no real customer data)."""
from __future__ import annotations

from tests.invoice_helpers import _workbook_bytes

OFFICE_ADDRESS = "보내는사람주소(발송금지/자연이랑·SK패밀리플러스마켓 사무실)"
EXPORT_COLUMNS = ["주문번호", "주문유형", "고객코드", "보내는사람", OFFICE_ADDRESS, "문의전화", "받는사람",
                  "받는사람주소", "우편번호", "전화번호", "휴대전화", "상품코드", "상품명", "주문수량", "판매단가",
                  "공급단가", "배송메시지", "전달메모"]
SQUID = "국내산 생물 갑오징어 1kg(4~8미 내외)"
CRAB = "속초 정품 자숙홍게 4kg(8~12미)"
OFFICE_MARK = "사무실-주소-사용금지"


def export_row(order_no: str = "O26100000001", name: str = "홍길동", **over) -> dict:
    row = {"주문번호": order_no, "주문유형": "일반", "고객코드": "C1", "보내는사람": "주문자",
           OFFICE_ADDRESS: OFFICE_MARK, "문의전화": "1588-0000", "받는사람": name,
           "받는사람주소": "경기도 가나시 다라구 마바로 1 (사아동* 자차아파트) 101동 202호", "우편번호": "13482  ",
           "전화번호": "", "휴대전화": "010-0000-0001", "상품코드": "2020100383", "상품명": SQUID,
           "주문수량": 1, "판매단가": 30000, "공급단가": 25000, "배송메시지": "문 앞에 놓아 주세요", "전달메모": ""}
    return {**row, **over}


def export_xlsx(*rows: dict) -> bytes:
    return _workbook_bytes("Sheet1", [EXPORT_COLUMNS] + [[r.get(c) for c in EXPORT_COLUMNS] for r in rows])
