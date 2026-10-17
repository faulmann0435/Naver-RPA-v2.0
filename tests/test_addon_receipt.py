"""Add-on rows (추가구성상품) follow their main product's form; dictionary rows show the order's 수령일."""
import pandas as pd

from core.config_loader import load_config
from core.dictionary import DictionaryEntry, DictionarySettings, ItemDictionary
from core.merger import merge_orders_with_dictionary
from core.option_key import receipt_date
from core.pipeline import follow_main_vendor, process_orders
from ui.option_logic import FormValues, row_fields, values_from_row
from store.rules_repo import DICTIONARY_COLUMNS
from store.validators import DATE_IN_TEXT_WARNING, validate_dictionary

EEL = "수령일 선택 (도착시간 지정불가): 10월 8일 수령 / 장어 선택: 손질장어 1kg"
ICE = "❄️ 아이스팩 추가: 아이스팩 1개 추가 (기본 2개 제공)"


# --- receipt_date ---------------------------------------------------------------------------

def test_receipt_date_from_option():
    assert receipt_date(EEL) == "10월 8일 수령"


def test_receipt_date_missing():
    assert receipt_date("장어 선택: 손질장어 1kg") == ""
    assert receipt_date(None) == ""
    assert receipt_date(float("nan")) == ""


# --- follow_main_vendor ---------------------------------------------------------------------

def _rows(*rows: tuple) -> pd.DataFrame:
    """rows: (상품주문번호, 묶음, 상품번호, 상품종류, _VendorID)"""
    return pd.DataFrame(rows, columns=["상품주문번호", "배송비 묶음번호", "상품번호", "상품종류", "_VendorID"])


MAIN = "조합형옵션상품"
ADDON = "추가구성상품"


def test_addon_takes_main_vendor_of_same_product():
    df = _rows(("1", "B1", "100", MAIN, "가리비"), ("2", "B1", "100", ADDON, "속초"))
    out = follow_main_vendor(df)
    assert out["_VendorID"].tolist() == ["가리비", "가리비"]
    assert df["_VendorID"].tolist() == ["가리비", "속초"]  # input not changed


def test_addon_without_main_keeps_own_vendor():
    df = _rows(("1", "B1", "100", ADDON, "속초"), ("2", "B1", "200", MAIN, "장어"))
    assert follow_main_vendor(df)["_VendorID"].tolist() == ["속초", "장어"]


def test_addon_does_not_cross_bundles():
    df = _rows(("1", "B1", "100", MAIN, "가리비"), ("2", "B2", "100", ADDON, "속초"))
    assert follow_main_vendor(df)["_VendorID"].tolist() == ["가리비", "속초"]


def test_addon_follows_first_main_when_main_rows_split():
    df = _rows(
        ("3", "B1", "100", MAIN, "킹선어"), ("2", "B1", "100", MAIN, "속초"),
        ("4", "B1", "100.0", ADDON, "X"),
    )
    assert follow_main_vendor(df)["_VendorID"].tolist() == ["킹선어", "속초", "속초"]


def test_addon_already_on_a_main_form_stays():
    df = _rows(
        ("2", "B1", "100", MAIN, "킹선어"), ("3", "B1", "100", MAIN, "속초"),
        ("4", "B1", "100", ADDON, "속초"),
    )
    assert follow_main_vendor(df)["_VendorID"].tolist() == ["킹선어", "속초", "속초"]


def test_no_kind_column_is_unchanged():
    df = _rows(("1", "B1", "100", MAIN, "가리비")).drop(columns=["상품종류"])
    assert follow_main_vendor(df).equals(df)


def _with_entries(df: pd.DataFrame, *own: bool | None) -> pd.DataFrame:
    """Attach a `_dict_entry` column: own_vendor flag per row, None = no dictionary match."""
    entries = [None if flag is None else DictionaryEntry("100", "k", "V", "t", own_vendor=flag) for flag in own]
    out = df.copy()
    out["_dict_entry"] = pd.Series(entries, index=df.index, dtype=object)
    return out


def test_own_vendor_addon_keeps_its_vendor():
    df = _with_entries(_rows(("1", "B1", "100", MAIN, "게"), ("2", "B1", "100", ADDON, "가리비")), False, True)
    assert follow_main_vendor(df)["_VendorID"].tolist() == ["게", "가리비"]


def test_unflagged_addon_still_follows():
    df = _with_entries(_rows(("1", "B1", "100", MAIN, "게"), ("2", "B1", "100", ADDON, "속초")), False, False)
    assert follow_main_vendor(df)["_VendorID"].tolist() == ["게", "게"]


def test_addon_without_dictionary_match_still_follows():
    df = _with_entries(_rows(("1", "B1", "100", MAIN, "게"), ("2", "B1", "100", ADDON, "속초")), None, None)
    assert follow_main_vendor(df)["_VendorID"].tolist() == ["게", "게"]


def test_own_vendor_addon_is_not_taken_as_a_main_row():
    df = _with_entries(
        _rows(("1", "B1", "100", MAIN, "게"), ("2", "B1", "100", ADDON, "가리비"), ("3", "B1", "100", ADDON, "속초")),
        False, True, False,
    )
    assert follow_main_vendor(df)["_VendorID"].tolist() == ["게", "가리비", "게"]


def test_pipeline_own_vendor_addon_goes_to_its_own_form():
    crab, scallop = "킹크랩 선택: 2kg", "추가상품: 참가리비 1kg"
    dictionary = _dictionary(
        {"option_key": crab, "vendor_id": "게 발주양식", "display_template": "킹크랩 {수량}마리"},
        {"option_key": scallop, "vendor_id": "가리비 발주양식", "display_template": "참가리비 {수량}kg", "own_vendor": "1"},
    )
    order = pd.DataFrame({
        "상품주문번호": ["A1", "A2"], "상품번호": ["777", "777"], "상품명": ["게", "게"],
        "상품종류": [MAIN, ADDON], "옵션정보": [crab, scallop], "수량": [1, 1],
        "배송비 묶음번호": ["B1", "B1"], "결제일": ["2026-10-06", "2026-10-06"],
    })
    merged = process_orders(order, load_config("config.xlsx"), dictionary).merged
    assert sorted(merged["_VendorID"].tolist()) == ["가리비 발주양식", "게 발주양식"]


# --- own_vendor in the dictionary entry and the option form --------------------------------------

def test_entry_own_vendor_parsing():
    def entry(**extra: str) -> DictionaryEntry:
        row = {**dict.fromkeys(DICTIONARY_COLUMNS, ""), "enabled": "1", "product_no": "1", "option_key": "k",
               "vendor_id": "V", "display_template": "t", **extra}
        found = ItemDictionary(pd.DataFrame([row], columns=DICTIONARY_COLUMNS)).lookup("1", "k", DictionarySettings())
        assert found is not None
        return found

    assert entry(own_vendor="1").own_vendor and entry(own_vendor="true").own_vendor
    assert not entry(own_vendor="0").own_vendor and not entry(own_vendor="").own_vendor
    old_csv = pd.DataFrame([{"enabled": "1", "product_no": "1", "option_key": "k", "vendor_id": "V"}])
    assert not ItemDictionary(old_csv).lookup("1", "k", DictionarySettings()).own_vendor  # column missing


def test_option_form_values_round_trip_own_vendor():
    values = values_from_row({"vendor_id": "V", "display_template": "t", "own_vendor": "1"})
    assert values.own_vendor and row_fields(values)["own_vendor"] == "1"
    assert not values_from_row({"vendor_id": "V"}).own_vendor
    assert row_fields(FormValues(vendor="V"))["own_vendor"] == "0"


# --- merge with receipt dates ---------------------------------------------------------------

def _entry(key: str, template: str = "", sum_group: str = "", unit: float | None = None) -> DictionaryEntry:
    return DictionaryEntry("1", key, "V", template, "", sum_group, unit, False)


def _merge(rows: list[tuple[DictionaryEntry, int, str]]) -> str:
    data = [{
        "배송비 묶음번호": "B1", "_VendorID": "V", "수량": qty, "상품명": "p", "옵션정보": "o",
        "processed_option": "", "_calculated_weight": 0.0, "_is_formatted": False,
        "배송메세지": "", "결제일": "2026-01-01", "_dict_entry": entry, "_receipt_date": date,
    } for entry, qty, date in rows]
    out = merge_orders_with_dictionary(pd.DataFrame(data), separator=", ")
    assert "_receipt_date" not in out.columns
    return str(out["processed_option"].iloc[0])


SCALLOP = _entry("s", sum_group="국산참가리비 (중~대)", unit=1.0)
PACK = _entry("i", template="아이스팩 {수량}개")


def test_date_prefixed_and_undated_addon_plain():
    assert _merge([(SCALLOP, 2, "10월 8일 수령"), (PACK, 1, "")]) == "10월 8일 수령 / 국산참가리비 (중~대) 2kg, 아이스팩 1개"


def test_same_date_weights_summed():
    assert _merge([(SCALLOP, 1, "10월 8일 수령"), (SCALLOP, 2, "10월 8일 수령")]) == "10월 8일 수령 / 국산참가리비 (중~대) 3kg"


def test_different_dates_kept_apart():
    assert _merge([(SCALLOP, 1, "10월 8일 수령"), (SCALLOP, 2, "10월 9일 수령")]) == (
        "10월 8일 수령 / 국산참가리비 (중~대) 1kg, 10월 9일 수령 / 국산참가리비 (중~대) 2kg"
    )


# --- whole pipeline -------------------------------------------------------------------------

def _dictionary(*rows: dict) -> ItemDictionary:
    base = {"channel": "naver", "enabled": "1", "product_no": "777"}
    frame = pd.DataFrame(
        [{**dict.fromkeys(DICTIONARY_COLUMNS, ""), **base, **r} for r in rows], columns=DICTIONARY_COLUMNS
    )
    return ItemDictionary(frame)


def test_pipeline_eel_with_ice_pack_is_one_eel_line():
    dictionary = _dictionary(
        {"option_key": EEL, "vendor_id": "장어 발주양식", "display_template": "손질장어 {수량}kg"},
        {"option_key": ICE, "vendor_id": "속초 발주양식", "display_template": "아이스팩 {수량}개"},
    )
    order = pd.DataFrame({
        "상품주문번호": ["A1", "A2"], "상품번호": ["777", "777"], "상품명": ["장어", "장어"],
        "상품종류": [MAIN, ADDON], "옵션정보": [EEL, ICE], "수량": [1, 1],
        "배송비 묶음번호": ["B1", "B1"], "결제일": ["2026-10-06", "2026-10-06"],
    })
    merged = process_orders(order, load_config("config.xlsx"), dictionary).merged
    assert merged["_VendorID"].tolist() == ["장어 발주양식"]
    assert merged["processed_option"].tolist() == ["10월 8일 수령 / 손질장어 1kg / 아이스팩 1개"]
    assert "_receipt_date" not in merged.columns


# --- validator warning ----------------------------------------------------------------------

def test_date_in_dictionary_text_warns():
    frame = pd.DataFrame(
        [{**dict.fromkeys(DICTIONARY_COLUMNS, ""), "enabled": "1", "product_no": "1", "vendor_id": "V1",
          "sum_group": "10월 16일 수령 / 국산참가리비", "unit_weight_kg": "1"}],
        columns=DICTIONARY_COLUMNS,
    )
    issues = validate_dictionary(frame, ["V1"], DictionarySettings())
    assert [(i.level, i.column, i.message) for i in issues] == [("warning", "sum_group", DATE_IN_TEXT_WARNING)]


def test_arrival_day_without_month_is_not_warned():
    frame = pd.DataFrame(
        [{**dict.fromkeys(DICTIONARY_COLUMNS, ""), "enabled": "1", "product_no": "1", "vendor_id": "V1",
          "display_template": "레드킹크랩 활어 (31일도착) {수량}마리"}],
        columns=DICTIONARY_COLUMNS,
    )
    assert validate_dictionary(frame, ["V1"], DictionarySettings()) == []
