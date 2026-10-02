"""Output of the pipeline must match the golden workbooks cell by cell (see regression_harness.py)."""
import pytest

from tests.regression_harness import compare_case, load_manifest

MANIFEST = load_manifest()

pytestmark = pytest.mark.skipif(
    not MANIFEST, reason="no golden data (run: python -m tests.regression_harness --update)"
)


@pytest.mark.parametrize("key", sorted(MANIFEST) or ["<none>"])
def test_output_matches_golden(key):
    problems = compare_case(key, MANIFEST[key])
    assert not problems, "\n".join(problems[:30])
