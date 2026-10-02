"""Golden comparison through process_orders: empty / unrelated dictionaries must not change any output."""
from io import BytesIO

import pandas as pd
import pytest

from core.config_loader import load_config
from core.dictionary import ItemDictionary
from core.loader import load_excel
from core.merger import filter_instruction_rows
from core.option_key import make_option_key
from core.pipeline import process_orders
from store.rules_repo import DICTIONARY_COLUMNS
from tests.regression_harness import (
    CONFIG_PATH,
    _Upload,
    compare_case,
    load_manifest,
    sample_dir,
)

MANIFEST = load_manifest()

pytestmark = pytest.mark.skipif(
    not MANIFEST, reason="no golden data (run: python -m tests.regression_harness --update)"
)


def _dictionary(*rows: dict) -> ItemDictionary:
    frame = pd.DataFrame([{**dict.fromkeys(DICTIONARY_COLUMNS, ""), **r} for r in rows], columns=DICTIONARY_COLUMNS)
    return ItemDictionary(frame)


def _runner(dictionary: ItemDictionary):
    return lambda df, config: process_orders(df, config, dictionary).files


UNRELATED = _dictionary(
    {"channel": "naver", "enabled": "1", "product_no": "1", "option_key": "없는 옵션: 값",
     "vendor_id": "ANY", "display_template": "NEVER {수량}"},
    {"channel": "naver", "enabled": "1", "product_no": "2", "option_key": "",
     "vendor_id": "ANY", "display_template": "NEVER"},
)


@pytest.mark.parametrize("key", sorted(MANIFEST) or ["<none>"])
def test_empty_dictionary_matches_golden(key):
    problems = compare_case(key, MANIFEST[key], runner=_runner(ItemDictionary.empty()))
    assert not problems, "\n".join(problems[:30])


@pytest.mark.parametrize("key", sorted(MANIFEST) or ["<none>"])
def test_unrelated_dictionary_matches_golden(key):
    problems = compare_case(key, MANIFEST[key], runner=_runner(UNRELATED))
    assert not problems, "\n".join(problems[:30])


def test_matching_entry_changes_output():
    key = min(MANIFEST)
    entry = MANIFEST[key]
    df = filter_instruction_rows(load_excel(_Upload(sample_dir() / entry["input"])))
    if "상품번호" not in df.columns or df.empty:
        pytest.skip("first golden case has no 상품번호 column")
    first = df.iloc[0]
    vendor = next(iter(entry["files"].values()))["vendor"]
    dictionary = _dictionary({
        "channel": "naver", "enabled": "1", "product_no": str(first["상품번호"]),
        "option_key": make_option_key(first.get("옵션정보")), "vendor_id": vendor,
        "display_template": "TEST {수량}",
    })
    result = process_orders(df, load_config(str(CONFIG_PATH)), dictionary)
    assert result.stats["rows_dictionary"] >= 1
    assert result.stats["rows_total"] == result.stats["rows_dictionary"] + result.stats["rows_auto"]
    texts = [
        str(cell)
        for f in result.files
        for cell in pd.read_excel(BytesIO(f["data"].getvalue()), dtype=object).to_numpy().ravel()
    ]
    assert any("TEST" in t for t in texts)
