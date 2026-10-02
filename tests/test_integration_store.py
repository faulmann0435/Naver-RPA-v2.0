"""
Integration tests against the real data repository (network).

Run only on demand:  RPA_INTEGRATION=1 python -m pytest tests/test_integration_store.py
Uses the [data_store] section of .streamlit/secrets.toml (branch is forced to "dev").
"""
import os
import uuid

import pytest
import tomllib

from store.base import ConflictError
from store.github_store import GitHubStore
from store.rules_repo import load_config_from_store
from tests.regression_harness import CONFIG_PATH, REPO_ROOT, compare_case, load_manifest

SECRETS = REPO_ROOT / ".streamlit" / "secrets.toml"

pytestmark = pytest.mark.skipif(
    os.environ.get("RPA_INTEGRATION") != "1" or not SECRETS.exists(),
    reason="set RPA_INTEGRATION=1 and provide .streamlit/secrets.toml",
)


@pytest.fixture(scope="module")
def dev_store() -> GitHubStore:
    with SECRETS.open("rb") as f:
        section = dict(tomllib.load(f)["data_store"])
    section["branch"] = "dev"
    return GitHubStore.from_secrets(section)


def test_dev_rules_reproduce_golden_outputs(dev_store):
    manifest = load_manifest()
    if not manifest:
        pytest.skip("no golden data")
    config = load_config_from_store(dev_store, str(CONFIG_PATH))
    failures = {k: compare_case(k, v, config=config) for k, v in manifest.items()}
    assert not any(failures.values()), {k: v[:5] for k, v in failures.items() if v}


def test_stale_sha_is_rejected_and_file_cleaned_up(dev_store):
    path = f"_it_{uuid.uuid4().hex[:8]}.txt"
    sha = dev_store.write_text(path, "a", None, "test: integration create")
    try:
        with pytest.raises(ConflictError):
            dev_store.write_text(path, "b", "0" * 40, "test: stale sha")
        assert dev_store.read_text(path).content == "a"
    finally:
        dev_store._session.delete(  # cleanup; DataStore has no delete on purpose
            dev_store._contents_url(path),
            headers=dev_store._headers,
            json={"message": "test: integration cleanup", "sha": sha, "branch": "dev"},
            timeout=15,
        )
