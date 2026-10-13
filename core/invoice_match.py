"""송장 매칭: match the courier's tracking-number file (한진 송장파일) to the Naver order rows.

Pure functions, no I/O beyond parsing xlsx bytes. Streamlit and the Telegram bot share this core.

    read_invoice_file(bytes)   -> list[Invoice]        one per box
    read_order_export(bytes)   -> DataFrame            the Naver order export (decrypt before)
    normalize_orders(df)       -> DataFrame            ids/zip/phones as text (also for the API table)
    match_invoices(df, inv)    -> MatchResult
    build_upload_workbook(...) -> bytes                the "발송처리" sheet Naver's 엑셀 일괄발송 reads
"""
from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from io import BytesIO

import openpyxl
import pandas as pd

KST = timezone(timedelta(hours=9))  # own copy: the Streamlit app ships without worker/
CARRIER_NAME = "한진택배"
DELIVERY_METHOD_LABEL = "택배,등기,소포"
UPLOAD_SHEET = "발송처리"
ORDER_SHEET = "발주발송관리"
STATUS_MATCHED, STATUS_CHECK = "MATCHED", "NEEDS_CHECK"
NOTE_NO_ORDER = "주문을 못 찾음"
READY_STATUS = "발송대기"
HEADER_SCAN_ROWS = 5

COL_ID, COL_ORDER_ID = "상품주문번호", "주문번호"
COL_METHOD, COL_CARRIER, COL_TRACKING = "배송방법", "택배사", "송장번호"
COL_NAME, COL_ZIP, COL_ADDRESS = "수취인명", "우편번호", "통합배송지"
COL_PHONE1, COL_PHONE2 = "수취인연락처1", "수취인연락처2"
COL_BUNDLE, COL_STATUS, COL_ITEM = "배송비 묶음번호", "주문상태", "상품명"
REQUIRED_ORDER_COLUMNS = (COL_ID, COL_NAME, COL_ZIP, COL_ADDRESS, COL_PHONE1)
TEXT_COLUMNS = (COL_ID, COL_ORDER_ID, COL_ZIP, COL_BUNDLE, "상품번호", "구매자연락처",
                COL_PHONE1, COL_PHONE2, COL_TRACKING)

HANJIN_TRACKING = "운송장번호"
HANJIN_REQUIRED = ("받으시는 분", "받는분우편번호", "받는분총주소")
# A purchase order the supplier sends back with the tracking numbers in a column named after the courier.
RETURNED_TRACKING, RETURNED_NAME = CARRIER_NAME, "받는사람"
RETURNED_REQUIRED = (RETURNED_NAME, "전화번호", "주소")
_BRACKET_TAIL = re.compile(r"\s*\[[^\]]*\]\s*$")


@dataclass(frozen=True)
class Invoice:
    """One box on the courier's file."""
    seq: str
    tracking: str
    name: str
    phone: str
    zip_code: str
    address: str
    item_name: str
    mobile: str = ""


@dataclass(frozen=True)
class Assignment:
    """A group of product orders (one shipment) and the tracking number(s) found for it."""
    product_order_ids: tuple[str, ...]
    tracking: str
    recipient: str
    items: str
    status: str
    includable: bool
    reason: str = ""


@dataclass(frozen=True)
class UnshippedOrder:
    """A shipment group no tracking number was found for."""
    product_order_ids: tuple[str, ...]
    recipient: str
    items: str


@dataclass(frozen=True)
class MatchResult:
    assignments: tuple[Assignment, ...]
    invoices_without_order: tuple[Invoice, ...]
    orders_without_invoice: tuple[UnshippedOrder, ...]
    invoices_without_tracking: tuple[Invoice, ...] = ()  # rows of the file with no tracking number

    def matched(self) -> tuple[Assignment, ...]:
        return tuple(a for a in self.assignments if a.status == STATUS_MATCHED)

    def needs_check(self) -> tuple[Assignment, ...]:
        return tuple(a for a in self.assignments if a.status == STATUS_CHECK)

    def includable_checks(self) -> tuple[Assignment, ...]:
        return tuple(a for a in self.needs_check() if a.includable)

    def includable(self) -> tuple[Assignment, ...]:
        return tuple(a for a in self.assignments if a.includable)


# ---------------------------------------------------------------- normalizing

def norm_name(value) -> str:
    return re.sub(r"\s+", "", _text(value))


def norm_phone(value) -> str:
    return re.sub(r"\D", "", _text(value))


def norm_zip(value) -> str:
    digits = re.sub(r"\D", "", _text(value))
    return digits.zfill(5) if digits else ""


def norm_order_address(value) -> str:
    return re.sub(r"\s+", "", _text(value))


def norm_invoice_address(value) -> str:
    """Hanjin appends ' [동, 지번]' to the address; drop it before comparing."""
    return norm_order_address(_BRACKET_TAIL.sub("", _text(value)))


def _blank(value) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _text(value) -> str:
    """Cell -> text; whole numbers lose the '.0', blanks become ''."""
    if _blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_orders(df: pd.DataFrame) -> pd.DataFrame:
    """Copy of the order table: id/zip/phone columns as text (zip zero-padded), blanks as ''."""
    out = df.reset_index(drop=True).astype(object)
    for column in out.columns:
        if column in TEXT_COLUMNS:
            out[column] = [_text(v) for v in out[column]]
        else:
            out[column] = [("" if _blank(v) else v) for v in out[column]]
    if COL_ZIP in out.columns:
        out[COL_ZIP] = [norm_zip(v) for v in out[COL_ZIP]]
    return out


# ---------------------------------------------------------------- reading files

def _open_sheet(data: bytes, preferred: str | None = None):
    if not data.startswith(b"PK"):
        raise ValueError("엑셀(.xlsx) 파일이 아닙니다. .xls 파일은 .xlsx로 저장한 뒤 올려 주세요.")
    try:
        workbook = openpyxl.load_workbook(BytesIO(data), data_only=True, read_only=True)
    except Exception as exc:  # any unreadable-file error becomes one readable message
        raise ValueError(f"엑셀 파일을 읽지 못했습니다: {exc}") from exc
    sheet = workbook[preferred] if preferred in workbook.sheetnames else workbook.worksheets[0]
    sheet.reset_dimensions()  # Naver's export declares a one-cell size; read-only mode would stop there
    return sheet


def _find_header(rows: list[tuple], key: str) -> int | None:
    for position, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        if any(_text(cell) == key for cell in row):
            return position
    return None


def read_invoice_file(data: bytes) -> list[Invoice]:
    """The courier's xlsx, or a purchase order a supplier sent back with a 한진택배 column.

    One Invoice per non-empty row; rows without a tracking number come back with tracking ''
    so match_invoices can report them (a supplier may have left a box out).
    """
    rows = list(_open_sheet(data).iter_rows(values_only=True))
    hanjin_at = _find_header(rows, HANJIN_TRACKING)
    if hanjin_at is not None:
        return _read_rows(rows, hanjin_at, HANJIN_REQUIRED, _hanjin_invoice)
    returned_at = _find_header(rows, RETURNED_TRACKING)
    if returned_at is not None and RETURNED_NAME in [_text(c) for c in rows[returned_at]]:
        return _read_rows(rows, returned_at, RETURNED_REQUIRED, _returned_invoice)
    raise ValueError("송장 파일이 아닙니다. 한진 송장파일('운송장번호' 열) 또는 "
                     "'한진택배' 열에 송장번호를 적은 발주서를 보내 주세요.")


def _read_rows(rows: list[tuple], header_at: int, required: tuple[str, ...], make) -> list[Invoice]:
    header = [_text(c) for c in rows[header_at]]
    missing = [c for c in required if c not in header]
    if missing:
        raise ValueError(f"송장 파일에 필요한 열이 없습니다: {', '.join(missing)}")
    invoices = []
    for line_no, row in enumerate(rows[header_at + 1:], start=header_at + 2):
        cells = dict(zip(header, row, strict=False))
        if any(_text(v) for v in cells.values()):
            invoices.append(make(cells, line_no))
    return invoices


def _hanjin_invoice(cells: dict, line_no: int) -> Invoice:
    return Invoice(
        seq=_text(cells.get("순번")) or str(line_no), tracking=_text(cells.get(HANJIN_TRACKING)),
        name=_text(cells.get("받으시는 분")), phone=_text(cells.get("받으시는 분 전화")),
        mobile=_text(cells.get("받는분핸드폰")), zip_code=norm_zip(cells.get("받는분우편번호")),
        address=_text(cells.get("받는분총주소")), item_name=_text(cells.get("품목명")))


def _returned_invoice(cells: dict, line_no: int) -> Invoice:
    """A purchase-order row (받는사람/전화번호/주소/구분); it has no zip code."""
    return Invoice(
        seq=str(line_no), tracking=_text(cells.get(RETURNED_TRACKING)), name=_text(cells.get(RETURNED_NAME)),
        phone=_text(cells.get("전화번호")), zip_code="", address=_text(cells.get("주소")),
        item_name=_text(cells.get("구분")))


def read_order_export(data: bytes) -> pd.DataFrame:
    """The Naver order export (already decrypted): every column, ids/zip/phones as text."""
    rows = list(_open_sheet(data, ORDER_SHEET).iter_rows(values_only=True))
    header_at = _find_header(rows, COL_ID)
    if header_at is None:
        raise ValueError("네이버 주문서가 아닙니다. '상품주문번호' 열을 찾지 못했습니다.")
    header = [_text(c) for c in rows[header_at]]
    keep = [i for i, name in enumerate(header) if name]
    body = [[(row[i] if i < len(row) else None) for i in keep] for row in rows[header_at + 1:]]
    body = [r for r in body if any(not _blank(v) for v in r)]
    return normalize_orders(pd.DataFrame(body, columns=[header[i] for i in keep]))


# ---------------------------------------------------------------- matching

@dataclass(frozen=True)
class _Group:
    ids: tuple[str, ...]
    recipient: str
    name: str
    zip_code: str
    address: str
    phones: frozenset[str]
    items: str


def _items_summary(names: list[str]) -> str:
    first = next((n for n in names if n), "")
    return first if len(names) <= 1 else f"{first} 외 {len(names) - 1}건"


def _eligible(orders: pd.DataFrame) -> pd.DataFrame:
    mask = pd.Series(True, index=orders.index)
    if COL_STATUS in orders.columns:
        mask &= orders[COL_STATUS].map(_text) == READY_STATUS
    if COL_TRACKING in orders.columns:
        mask &= orders[COL_TRACKING].map(_text) == ""
    return orders[mask]


def _group_orders(orders: pd.DataFrame) -> list[_Group]:
    missing = [c for c in REQUIRED_ORDER_COLUMNS if c not in orders.columns]
    if missing:
        raise ValueError(f"주문서에 필요한 열이 없습니다: {', '.join(missing)}")
    buckets: dict[tuple, list[dict]] = defaultdict(list)
    for row in _eligible(orders).to_dict("records"):
        bundle = _text(row.get(COL_BUNDLE))
        key = ("B", bundle) if bundle else (
            "N", norm_name(row[COL_NAME]), norm_zip(row[COL_ZIP]), norm_order_address(row[COL_ADDRESS]))
        buckets[key].append(row)
    return [_group(rows) for rows in buckets.values()]


def _group(rows: list[dict]) -> _Group:
    first = rows[0]
    phones = {norm_phone(r.get(c)) for r in rows for c in (COL_PHONE1, COL_PHONE2)}
    return _Group(
        ids=tuple(_text(r[COL_ID]) for r in rows), recipient=_text(first[COL_NAME]),
        name=norm_name(first[COL_NAME]), zip_code=norm_zip(first[COL_ZIP]),
        address=norm_order_address(first[COL_ADDRESS]), phones=frozenset(p for p in phones if p),
        items=_items_summary([_text(r.get(COL_ITEM)) for r in rows]))


def _mismatch(group: _Group, invoice: Invoice) -> tuple[bool, bool]:
    """(phone_ok, address_ok)"""
    phones = {norm_phone(invoice.phone), norm_phone(invoice.mobile)} - {""}
    return bool(phones & group.phones), norm_invoice_address(invoice.address) == group.address


def _is_candidate(group: _Group, invoice: Invoice) -> bool:
    """Same name and zip; a file without zip codes needs the phone or the address to agree as well."""
    if norm_name(invoice.name) != group.name:
        return False
    zip_code = norm_zip(invoice.zip_code)
    if zip_code:
        return zip_code == group.zip_code
    return any(_mismatch(group, invoice))


def _problem_reason(phone_ok: bool, addr_ok: bool) -> str:
    return "전화번호·주소 다름" if not phone_ok and not addr_ok else (
        "전화번호 다름" if not phone_ok else "주소 다름")


def _assignment(group: _Group, tracking: str, status: str, includable: bool, reason: str = "") -> Assignment:
    return Assignment(group.ids, tracking, group.recipient, group.items, status, includable, reason)


def _decide(group: _Group, picked: list[int], shared_by: list[int], invoices: list[Invoice],
            groups: list[_Group], candidates: list[list[int]]) -> Assignment:
    if len(picked) > 1:
        tracking = ", ".join(invoices[i].tracking for i in picked)
        return _assignment(group, tracking, STATUS_CHECK, False, "송장 여러 개")
    invoice = invoices[picked[0]]
    phone_ok, addr_ok = _mismatch(group, invoice)
    if len(shared_by) == 1:
        if phone_ok and addr_ok:
            return _assignment(group, invoice.tracking, STATUS_MATCHED, True)
        return _assignment(group, invoice.tracking, STATUS_CHECK, True, _problem_reason(phone_ok, addr_ok))
    all_clean = all(len(candidates[g]) == 1 and all(_mismatch(groups[g], invoice))
                    for g in shared_by)
    if all_clean:
        return _assignment(group, invoice.tracking, STATUS_CHECK, True, "따로 주문했지만 한 상자로 보낸 것 같음")
    return _assignment(group, invoice.tracking, STATUS_CHECK, False, "여러 주문에 걸친 송장")


def match_invoices(orders: pd.DataFrame, invoices: list[Invoice]) -> MatchResult:
    """Pair each shipment group of the order table with the tracking number(s) of the same recipient."""
    no_tracking = tuple(i for i in invoices if not i.tracking)
    invoices = [i for i in invoices if i.tracking]
    groups = _group_orders(orders)
    candidates = [[i for i, inv in enumerate(invoices) if _is_candidate(g, inv)] for g in groups]
    sharers: dict[int, list[int]] = defaultdict(list)
    for g, picked in enumerate(candidates):
        for i in picked:
            sharers[i].append(g)
    assignments, unshipped = [], []
    for g, group in enumerate(groups):
        if not candidates[g]:
            unshipped.append(UnshippedOrder(group.ids, group.recipient, group.items))
            continue
        shared = sharers[candidates[g][0]]
        assignments.append(_decide(group, candidates[g], shared, invoices, groups, candidates))
    leftovers = tuple(inv for i, inv in enumerate(invoices) if i not in sharers)
    return MatchResult(tuple(assignments), leftovers, tuple(unshipped), no_tracking)


# ---------------------------------------------------------------- upload workbook

def _upload_columns(columns: list[str]) -> list[str]:
    """All order columns; the three dispatch columns are added after the order number when absent."""
    result = list(columns)
    anchor = result.index(COL_ORDER_ID) + 1 if COL_ORDER_ID in result else (
        result.index(COL_ID) + 1 if COL_ID in result else 0)
    for name in (COL_METHOD, COL_CARRIER, COL_TRACKING):
        if name not in result:
            result.insert(anchor, name)
            anchor += 1
    return result


def _excel_value(value):
    if _blank(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if hasattr(value, "item") and not isinstance(value, datetime):
        return value.item()
    return value


def build_upload_workbook(orders: pd.DataFrame, assignments, carrier: str = CARRIER_NAME,
                          method: str = DELIVERY_METHOD_LABEL) -> bytes:
    """xlsx for Naver's 엑셀 일괄발송: sheet 발송처리, header in row 1, only the assigned rows."""
    tracking_of = {pid: a.tracking for a in assignments for pid in a.product_order_ids}
    table = normalize_orders(orders)
    columns = _upload_columns(list(table.columns))
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = UPLOAD_SHEET
    sheet.append(columns)
    for row in table.to_dict("records"):
        tracking = tracking_of.get(_text(row.get(COL_ID)))
        if tracking is None:
            continue
        row = {**row, COL_METHOD: method, COL_CARRIER: carrier, COL_TRACKING: tracking}
        sheet.append([_excel_cell(row.get(c, ""), c) for c in columns])
    for position, name in enumerate(columns, start=1):
        if name in TEXT_COLUMNS:
            for row_number in range(2, sheet.max_row + 1):
                sheet.cell(row_number, position).number_format = "@"
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _excel_cell(value, column: str):
    if column in TEXT_COLUMNS or column in (COL_METHOD, COL_CARRIER):
        return _text(value)
    return _excel_value(value)
