"""RPA_DATA_DIR: env var, then .env, then the folder next to the repo."""
from __future__ import annotations

from worker import paths


def test_default_is_repo_parent(monkeypatch, tmp_path):
    monkeypatch.delenv(paths.DATA_DIR_KEY, raising=False)
    assert paths.data_dir(tmp_path / "missing.env") == paths.REPO_ROOT.parent


def test_env_var_wins_over_env_file(monkeypatch, tmp_path):
    env_file = tmp_path / ".env"
    env_file.write_text(f"{paths.DATA_DIR_KEY}={tmp_path / 'fromfile'}\n", encoding="utf-8")
    monkeypatch.setenv(paths.DATA_DIR_KEY, str(tmp_path / "fromenv"))
    assert paths.data_dir(env_file) == tmp_path / "fromenv"


def test_env_file_used_when_env_var_absent(monkeypatch, tmp_path):
    env_file = tmp_path / ".env"
    env_file.write_text(f"{paths.DATA_DIR_KEY}={tmp_path / 'fromfile'}\n", encoding="utf-8")
    monkeypatch.delenv(paths.DATA_DIR_KEY, raising=False)
    assert paths.data_dir(env_file) == tmp_path / "fromfile"
