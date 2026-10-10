"""Stand-in for the Naver API until the store's API key exists.

Turns an order Excel export back into the shape the API's product-order query returns, so the
adapter and the whole API path can be run and compared with the Excel path on real past orders.
It is the inverse of naver_adapter.FIELD_SOURCES; it proves the wiring, not the API's field names.
"""
from __future__ import annotations

from collections.abc import Mapping

import pandas as pd

from worker.naver_adapter import (
    ADDRESS_COLUMN,
    FIELD_SOURCES,
    STATUS_COLUMN,
    STATUS_LABELS,
)

ZIP_DIGITS = 5
_STATUS_CODES = {label: code for code, label in STATUS_LABELS.items()}


def _api_value(column: str, value):
    """Excel cell -> the JSON value the API would send (strings for ids, ISO times, None for blanks)."""
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.isoformat(timespec="milliseconds") + "+09:00"
    if column == "우편번호":
        return str(int(value)).zfill(ZIP_DIGITS)
    if column == "수량":
        return int(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _set_path(target: dict, path: str, value) -> dict:
    """Return a copy of `target` with `value` at the dotted `path`."""
    head, _, rest = path.partition(".")
    if not rest:
        return {**target, head: value}
    return {**target, head: _set_path(target.get(head) or {}, rest, value)}


def row_to_item(row: Mapping) -> dict:
    """One Excel row -> one API item {order, productOrder}."""
    item: dict = {"order": {}, "productOrder": {}}
    for column, (obj_name, path) in FIELD_SOURCES.items():
        if column in row:
            item = {**item, obj_name: _set_path(item[obj_name], path, _api_value(column, row[column]))}
    address = _api_value(ADDRESS_COLUMN, row.get(ADDRESS_COLUMN))
    item = {**item, "productOrder": _set_path(item["productOrder"], "shippingAddress.baseAddress", address)}
    status = _api_value(STATUS_COLUMN, row.get(STATUS_COLUMN))
    code = _STATUS_CODES.get(status, status)
    return {**item, "productOrder": {**item["productOrder"], "productOrderStatus": code}}


def excel_to_api_items(frame: pd.DataFrame) -> list[dict]:
    """Excel export rows -> API `data` list."""
    return [row_to_item(row) for row in frame.to_dict(orient="records")]
