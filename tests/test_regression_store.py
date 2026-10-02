"""Golden comparison with the rules coming from the (in-memory) data store instead of config.xlsx."""
import pytest

from core.config_loader import read_config_sheets
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
def store_config() -> dict:
    store = MemoryStore()
    files = sheets_to_rule_csvs(read_config_sheets(str(CONFIG_PATH)), "test", "2026-10-02T00:00:00+09:00")
    for name, text in files.items():
        store.write_text(name, text, None, "seed")
    return load_config_from_store(store, str(CONFIG_PATH))


@pytest.mark.parametrize("key", sorted(MANIFEST) or ["<none>"])
def test_store_output_matches_golden(key, store_config):
    if not case_available(MANIFEST[key]):
        pytest.skip("encrypted sample: set RPA_ORDER_PASSWORD")
    problems = compare_case(key, MANIFEST[key], config=store_config)
    assert not problems, "\n".join(problems[:30])
