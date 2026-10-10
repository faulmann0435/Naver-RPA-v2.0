"""Naver API order details -> Excel-shaped order table."""
import pandas as pd
import pytest

from worker.naver_adapter import (
    COLUMN_ORDER,
    AdapterError,
    details_to_dataframe,
    split_shippable,
)
from worker.simulator import excel_to_api_items


def _item(**product_order):
    base = {
        "order": {
            "orderId": "2026012157745901",
            "ordererName": "김구매",
            "ordererTel": "010-1111-2222",
            "ordererId": "buyer***",
            "paymentDate": "2026-01-21T16:27:55.0+09:00",
            "orderDate": "2026-01-21T16:27:40.0+09:00",
        },
        "productOrder": {
            "productOrderId": "2026012151232691",
            "productName": "속초 오징어순대",
            "productOption": "순대 선택: 야채 오징어순대 2미",
            "quantity": 2,
            "packageNumber": "2026012128585955",
            "productId": "5568579375",
            "productClass": "조합형옵션상품",
            "productOrderStatus": "PAYED",
            "totalPaymentAmount": 31800,
            "shippingAddress": {
                "name": "박수취",
                "tel1": "010-3333-4444",
                "zipCode": "04524",
                "baseAddress": "서울특별시 중구 세종대로 110",
                "detailedAddress": "3층",
            },
        },
    }
    return {**base, "productOrder": {**base["productOrder"], **product_order}}


def test_columns_follow_excel_export_order():
    df = details_to_dataframe([_item()])
    assert list(df.columns) == list(COLUMN_ORDER)


def test_ids_and_amounts_are_int64_like_the_excel_export():
    df = details_to_dataframe([_item()])
    for column in ("상품주문번호", "주문번호", "우편번호", "수량", "배송비 묶음번호", "상품번호"):
        assert df[column].dtype == "int64", column
    assert df.loc[0, "상품번호"] == 5568579375
    assert df.loc[0, "우편번호"] == 4524  # the Excel export also drops the leading zero


def test_address_joins_base_and_detail():
    assert details_to_dataframe([_item()]).loc[0, "통합배송지"] == "서울특별시 중구 세종대로 110 3층"


def test_payment_time_is_korean_local_time_without_zone():
    df = details_to_dataframe([_item()])
    assert df.loc[0, "결제일"] == pd.Timestamp("2026-01-21 16:27:55")
    utc = _item()
    utc = {**utc, "order": {**utc["order"], "paymentDate": "2026-01-21T07:27:55Z"}}
    assert details_to_dataframe([utc]).loc[0, "결제일"] == pd.Timestamp("2026-01-21 16:27:55")


def test_blank_values_become_nan_like_empty_excel_cells():
    df = details_to_dataframe([_item(shippingMemo="  ")])
    assert pd.isna(df.loc[0, "배송메세지"])
    assert pd.isna(df.loc[0, "수취인연락처2"])


def test_status_code_is_shown_in_korean():
    assert details_to_dataframe([_item()]).loc[0, "주문상태"] == "발송대기"


def test_missing_product_order_id_is_refused():
    with pytest.raises(AdapterError):
        details_to_dataframe([_item(productOrderId=None)])


def test_non_numeric_id_is_refused_with_column_name():
    with pytest.raises(AdapterError, match="상품번호"):
        details_to_dataframe([_item(productId="abc")])


def test_empty_response_gives_empty_table_with_columns():
    df = details_to_dataframe([])
    assert df.empty and list(df.columns) == list(COLUMN_ORDER)


def test_input_items_are_not_changed():
    item = _item()
    before = repr(item)
    details_to_dataframe([item])
    assert repr(item) == before


def test_simulator_round_trip_keeps_engine_columns():
    original = details_to_dataframe([_item(), _item(productOrderId="2026012151232692", quantity=1)])
    again = details_to_dataframe(excel_to_api_items(original))
    pd.testing.assert_frame_equal(original, again)


def test_payment_time_drops_milliseconds_like_the_excel_export():
    item = _item()
    item = {**item, "order": {**item["order"], "paymentDate": "2026-01-21T16:27:55.873+09:00"}}
    assert details_to_dataframe([item]).loc[0, "결제일"] == pd.Timestamp("2026-01-21 16:27:55")


def test_only_paid_orders_without_claims_are_shipped():
    paid = _item()
    canceled = _item(productOrderId="2", productOrderStatus="CANCELED", claimType="CANCEL")
    cancel_requested = _item(productOrderId="3", claimType="CANCEL", claimStatus="CANCEL_REQUEST")
    shipped = _item(productOrderId="4", productOrderStatus="DELIVERING")
    keep, left_out = split_shippable([paid, canceled, cancel_requested, shipped])
    assert keep == [paid]
    assert left_out == [canceled, cancel_requested, shipped]
