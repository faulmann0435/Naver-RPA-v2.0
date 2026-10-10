"""The API path (Excel -> simulated API JSON -> adapter -> engine) must give the golden workbooks."""
import pytest

from core.merger import filter_instruction_rows
from core.pipeline import process_all_data
from tests.regression_harness import case_available, compare_case, load_manifest
from worker.naver_adapter import details_to_dataframe
from worker.simulator import excel_to_api_items

MANIFEST = load_manifest()

pytestmark = pytest.mark.skipif(
    not MANIFEST, reason="no golden data (run: python -m tests.regression_harness --update)"
)


def api_path_runner(df, config):
    items = excel_to_api_items(filter_instruction_rows(df))
    return process_all_data(details_to_dataframe(items), config)


@pytest.mark.parametrize("key", sorted(MANIFEST) or ["<none>"])
def test_api_path_matches_golden(key):
    if not case_available(MANIFEST[key]):
        pytest.skip("encrypted sample: set RPA_ORDER_PASSWORD")
    problems = compare_case(key, MANIFEST[key], runner=api_path_runner)
    assert not problems, "\n".join(problems[:30])
