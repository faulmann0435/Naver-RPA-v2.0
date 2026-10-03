"""Golden comparison with rules AND the purchase-order forms coming from the in-memory data store."""
import pytest

from core.config_loader import read_config_sheets
from store.layout_repo import OUTPUT_LAYOUT_FILE, layout_sheet_to_csv
from store.memory_store import MemoryStore
from store.rules_repo import load_config_from_store, sheets_to_rule_csvs
from tests.regression_harness import (
    CONFIG_PATH,
    case_available,
    compare_case,
    load_manifest,
)

MANIFEST = load_manifest()

pytestmark = pytest.mark.skipif(
    not MANIFEST, reason="no golden data (run: python -m tests.regression_harness --update)"
)


@pytest.fixture(scope="module")
def layout_store_config() -> dict:
    store = MemoryStore()
    sheets = read_config_sheets(str(CONFIG_PATH))
    for name, text in sheets_to_rule_csvs(sheets, "test", "2026-10-03T00:00:00+09:00").items():
        store.write_text(name, text, None, "seed")
    layout_text = layout_sheet_to_csv(sheets["OutputLayout"], "test", "2026-10-03T00:00:00+09:00")
    store.write_text(OUTPUT_LAYOUT_FILE, layout_text, None, "seed layout")
    return load_config_from_store(store, "this-file-does-not-exist.xlsx")


@pytest.mark.parametrize("key", sorted(MANIFEST) or ["<none>"])
def test_layout_store_output_matches_golden(key, layout_store_config):
    if not case_available(MANIFEST[key]):
        pytest.skip("encrypted sample: set RPA_ORDER_PASSWORD")
    problems = compare_case(key, MANIFEST[key], config=layout_store_config)
    assert not problems, "\n".join(problems[:30])
