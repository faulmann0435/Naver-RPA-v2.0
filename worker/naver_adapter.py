"""Naver Commerce API order details -> the same table the Naver order Excel export gives.

The v15 engine (core.pipeline.process_orders) reads the Excel export. This adapter turns the
response of POST /v1/pay-order/seller/product-orders/query (its `data` list) into a DataFrame
with the Excel column names and dtypes, so the engine, dictionary and layouts run unchanged.

Field names follow the API as reported by public integrations (see FIELD_SOURCES); confirm them
against the first real API response before production use.
"""
from __future__ import annotations

from collections.abc import Iterable, Mapping
from datetime import datetime, timedelta, timezone

import numpy as np
import pandas as pd

KST = timezone(timedelta(hours=9))

# Excel column -> (API object, field path). Object is "order" or "productOrder";
# a path with a dot reads inside productOrder.shippingAddress.
FIELD_SOURCES: dict[str, tuple[str, str]] = {
    "상품주문번호": ("productOrder", "productOrderId"),
    "주문번호": ("order", "orderId"),
    "구매자명": ("order", "ordererName"),
    "수취인명": ("productOrder", "shippingAddress.name"),
    "구매자연락처": ("order", "ordererTel"),
    "우편번호": ("productOrder", "shippingAddress.zipCode"),
    "상품명": ("productOrder", "productName"),
    "옵션정보": ("productOrder", "productOption"),
    "수량": ("productOrder", "quantity"),
    "수취인연락처1": ("productOrder", "shippingAddress.tel1"),
    "배송메세지": ("productOrder", "shippingMemo"),
    "구매자ID": ("order", "ordererId"),
    "배송비 묶음번호": ("productOrder", "packageNumber"),
    "결제일": ("order", "paymentDate"),
    "상품번호": ("productOrder", "productId"),
    "상품종류": ("productOrder", "productClass"),
    "최종 상품별 총 주문금액": ("productOrder", "totalPaymentAmount"),
    "수취인연락처2": ("productOrder", "shippingAddress.tel2"),
    "주문일시": ("order", "orderDate"),
}
ADDRESS_COLUMN = "통합배송지"
STATUS_COLUMN = "주문상태"

# The Excel export stores these as whole numbers (int64); the API sends them as strings.
INTEGER_COLUMNS = ("상품주문번호", "주문번호", "우편번호", "수량", "배송비 묶음번호", "상품번호",
                   "최종 상품별 총 주문금액")
DATETIME_COLUMNS = ("결제일", "주문일시")

# Column order of the Excel export, for the columns this adapter fills.
COLUMN_ORDER = (
    "상품주문번호", "주문번호", "구매자명", "수취인명", "구매자연락처", "우편번호", "상품명", "옵션정보",
    "수량", "수취인연락처1", ADDRESS_COLUMN, "배송메세지", "구매자ID", STATUS_COLUMN, "배송비 묶음번호",
    "결제일", "상품번호", "상품종류", "최종 상품별 총 주문금액", "수취인연락처2", "주문일시",
)

STATUS_LABELS: dict[str, str] = {
    "PAYMENT_WAITING": "결제대기",
    "PAYED": "결제완료",
    "DELIVERING": "배송중",
    "DELIVERED": "배송완료",
    "PURCHASE_DECIDED": "구매확정",
    "EXCHANGED": "교환",
    "CANCELED": "취소",
    "RETURNED": "반품",
    "CANCELED_BY_NOPAYMENT": "미결제취소",
}


class AdapterError(ValueError):
    """An API item is missing something the order table cannot do without."""


def _blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _read(item: Mapping, source: tuple[str, str]):
    obj_name, path = source
    current = item.get(obj_name) or {}
    for key in path.split("."):
        if not isinstance(current, Mapping):
            return None
        current = current.get(key)
    return None if _blank(current) else current


def _address(item: Mapping) -> str | None:
    shipping = (item.get("productOrder") or {}).get("shippingAddress") or {}
    parts = [str(shipping.get(k)).strip() for k in ("baseAddress", "detailedAddress") if not _blank(shipping.get(k))]
    return " ".join(parts) or None


def _status(item: Mapping) -> str | None:
    code = _read(item, ("productOrder", "productOrderStatus"))
    return None if code is None else STATUS_LABELS.get(str(code), str(code))


def item_to_row(item: Mapping) -> dict:
    """One API item ({order, productOrder, ...}) -> one Excel-shaped row (missing values are None)."""
    row = {column: _read(item, source) for column, source in FIELD_SOURCES.items()}
    if row["상품주문번호"] is None:
        raise AdapterError("상품주문번호(productOrderId)가 없는 주문이 있습니다.")
    return {**row, ADDRESS_COLUMN: _address(item), STATUS_COLUMN: _status(item)}


def _kst_naive(value) -> datetime | None:
    """ISO-8601 time from the API -> Korean local time without zone, as the Excel export shows it."""
    if value is None:
        return None
    parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(KST).replace(tzinfo=None)
    return parsed


def _to_int(column: str, value) -> int:
    try:
        return int(str(value).strip())
    except ValueError as exc:
        raise AdapterError(f"'{column}' 값이 숫자가 아닙니다: {value!r}") from exc


def _typed(frame: pd.DataFrame) -> pd.DataFrame:
    """Apply the Excel export's dtypes: ids/amounts int64, times datetime64, blanks NaN."""
    columns: dict[str, pd.Series] = {}
    for name in frame.columns:
        values = frame[name]
        if name in INTEGER_COLUMNS and values.notna().all():
            columns[name] = pd.Series([_to_int(name, v) for v in values], index=frame.index, dtype="int64")
        elif name in INTEGER_COLUMNS:
            columns[name] = pd.Series([np.nan if v is None else _to_int(name, v) for v in values],
                                      index=frame.index, dtype="float64")
        elif name in DATETIME_COLUMNS:
            columns[name] = pd.to_datetime(pd.Series([_kst_naive(v) for v in values], index=frame.index))
        else:
            columns[name] = values.map(lambda v: np.nan if v is None else v).astype(object)
    return pd.DataFrame(columns, index=frame.index)


def details_to_dataframe(items: Iterable[Mapping]) -> pd.DataFrame:
    """API `data` list -> order table with the Excel export's column names and dtypes."""
    rows = [item_to_row(item) for item in items]
    frame = pd.DataFrame(rows, columns=list(COLUMN_ORDER)) if rows else pd.DataFrame(columns=list(COLUMN_ORDER))
    return _typed(frame) if rows else frame
