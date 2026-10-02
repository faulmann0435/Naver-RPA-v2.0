import pandas as pd

from core.dictionary import DictionaryEntry
from core.merger import merge_orders_with_dictionary


def _entry(
    key: str = "k", template: str = "", qty1: str = "", sum_group: str = "",
    unit: float | None = None, append: bool = False, vendor: str = "V",
) -> DictionaryEntry:
    return DictionaryEntry("1", key, vendor, template, qty1, sum_group, unit, append)


def _frame(rows: list[tuple[DictionaryEntry | None, int]], auto_text: str = "AUTO") -> pd.DataFrame:
    data = []
    for i, (entry, qty) in enumerate(rows):
        data.append({
            "배송비 묶음번호": "B1", "_VendorID": entry.vendor_id if entry else "V", "수량": qty,
            "상품명": "p", "옵션정보": "o", "processed_option": auto_text if entry is None else "ignored",
            "_calculated_weight": 0.0 if entry is None else 9.9, "_is_formatted": False,
            "배송메세지": "", "결제일": f"2026-01-0{i + 1}", "_dict_entry": entry,
        })
    return pd.DataFrame(data)


def _option(rows: list[tuple[DictionaryEntry | None, int]]) -> str:
    out = merge_orders_with_dictionary(_frame(rows))
    assert len(out) == 1
    return str(out["processed_option"].iloc[0])


KODARI = _entry(template="속초 프리미엄코다리 {수량*8}마리, 소스 {수량}개")
A = _entry(key="a", sum_group="★제수용 피문어", unit=1.0)
B = _entry(key="b", sum_group="★제수용 피문어", unit=0.5)
C = _entry(key="c", template="(자숙)", append=True)
D = _entry(key="d", sum_group="동해 피문어", unit=1.0)


def test_1_template_multiplied():
    assert _option([(KODARI, 2)]) == "속초 프리미엄코다리 16마리, 소스 2개"


def test_2_same_entry_rows_are_summed():
    assert _option([(KODARI, 1), (KODARI, 1)]) == "속초 프리미엄코다리 16마리, 소스 2개"


def test_3_sum_group_with_suffix():
    assert _option([(A, 1), (B, 1), (C, 1)]) == "★제수용 피문어 1.5kg (자숙)"


def test_4_sum_group_integer_kg():
    assert _option([(D, 3), (C, 1)]) == "동해 피문어 3kg (자숙)"


def test_5_qty1_template():
    sea = _entry(template="깐멍게 500gx{수량}개", qty1="깐멍게 500g")
    assert _option([(sea, 1)]) == "깐멍게 500g"
    assert _option([(sea, 2)]) == "깐멍게 500gx2개"
    assert _option([(_entry(template="깐멍게 500gx{수량}개"), 1)]) == "깐멍게 500gx1개"


def test_6_plain_template_gets_qty_suffix():
    assert _option([(_entry(template="깐멍게 500g"), 2)]) == "깐멍게 500g (x2)"


def test_7_sum_group_under_one_kg_in_grams():
    assert _option([(_entry(sum_group="G", unit=0.5), 1)]) == "G 500g"


def test_8_mixed_bundle_dictionary_then_auto():
    assert _option([(KODARI, 1), (None, 1)]) == "속초 프리미엄코다리 8마리, 소스 1개 / AUTO"


def test_9_vendor_override_splits_bundle():
    other = _entry(template="X", vendor="OTHER")
    out = merge_orders_with_dictionary(_frame([(KODARI, 1), (other, 1)]))
    assert sorted(out["_VendorID"]) == ["OTHER", "V"]


def test_output_has_no_helper_column_and_auto_weight_only():
    out = merge_orders_with_dictionary(_frame([(KODARI, 1), (None, 1)]))
    assert "_dict_entry" not in out.columns
    assert out["_calculated_weight"].iloc[0] == 0.0
    assert out["수량"].iloc[0] == 2


def test_custom_separator_and_sum_interleaving():
    plain = _entry(key="p", template="P")
    out = merge_orders_with_dictionary(_frame([(A, 1), (plain, 1), (B, 1)]), separator=" | ")
    assert out["processed_option"].iloc[0] == "★제수용 피문어 1.5kg | P"
