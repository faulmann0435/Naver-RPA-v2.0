"""📦 주문처리 page with a 자연이랑 출고작업관리 Excel: detection, conversion, renamed files, registered 양식."""
from io import BytesIO

import openpyxl
import pandas as pd

from core.config_loader import load_config
from core.dictionary import DictionarySettings, ItemDictionary
from core.pipeline import process_orders
from core.withnature import CHANNEL as WITHNATURE
from store.rules_repo import DICTIONARY_COLUMNS
from tests.invoice_helpers import order_export_xlsx, order_row
from tests.withnature_helpers import CRAB, export_row, export_xlsx
from ui.order_logic import channel_files, withnature_orders

SQUID_FORM = "생물오징어, 갑오징어, 호래기, 실속과메기 발주양식"
SOKCHO_FORM = "속초 발주양식"
CRAB_NO = "자연이랑-2020100263"


def _config():
    return load_config("config.xlsx", password="1111")


def _orders():
    return withnature_orders(export_xlsx(
        export_row("O1", "김가나"), export_row("O2", "이다라", 상품코드="2020100263", 상품명=CRAB)))


def _vendor_of(result, name: str) -> str:
    merged = result.merged
    return str(merged.loc[merged["수취인명"] == name, "_VendorID"].iloc[0]).strip()


def test_withnature_file_becomes_the_naver_shaped_table():
    frame = _orders()
    assert list(frame["상품번호"]) == ["자연이랑-2020100383", CRAB_NO]
    assert {"수량", "상품명", "옵션정보", "배송비 묶음번호"} <= set(frame.columns)


def test_other_files_are_not_withnature():
    assert withnature_orders(order_export_xlsx(order_row("1"))) is None
    assert withnature_orders(b"junk") is None


def test_files_get_the_channel_prefix_only_for_withnature():
    files = [{"vendor": "A", "filename": "속초 발주양식_20261109.xlsx"}]
    assert channel_files(files, WITHNATURE)[0]["filename"] == "자연이랑_속초 발주양식_20261109.xlsx"
    assert channel_files(files, None) == files
    assert files[0]["filename"] == "속초 발주양식_20261109.xlsx"  # the original list is not changed


def test_unregistered_items_are_offered_for_registration():
    result = process_orders(_orders(), _config(), ItemDictionary.empty(), DictionarySettings())
    assert set(result.unmatched["product_no"].astype(str)) == {"자연이랑-2020100383", CRAB_NO}
    assert _vendor_of(result, "이다라") == SOKCHO_FORM  # 홍게 auto rule, before registration


def test_registered_form_wins_over_the_auto_rule():
    unmatched = process_orders(_orders(), _config(), ItemDictionary.empty(), DictionarySettings()).unmatched
    key = str(unmatched.loc[unmatched["product_no"].astype(str) == CRAB_NO, "option_key"].iloc[0])
    entry = {**dict.fromkeys(DICTIONARY_COLUMNS, ""), "channel": "naver", "enabled": "1", "product_no": CRAB_NO,
             "option_key": key, "vendor_id": SQUID_FORM, "display_template": "박달홍게 {수량}박스"}
    dictionary = ItemDictionary(pd.DataFrame([entry], columns=DICTIONARY_COLUMNS))
    result = process_orders(_orders(), _config(), dictionary, DictionarySettings())
    assert _vendor_of(result, "이다라") == SQUID_FORM
    assert CRAB_NO not in set(result.unmatched["product_no"].astype(str))
    squid_file = next(f for f in result.files if f["vendor"] == SQUID_FORM)
    flat = "\n".join(str(c) for r in openpyxl.load_workbook(BytesIO(squid_file["data"].getvalue())).active
                     .iter_rows(values_only=True) for c in r)
    assert "박달홍게 1박스" in flat and "이다라" in flat
