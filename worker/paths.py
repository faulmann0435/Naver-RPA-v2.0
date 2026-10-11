"""Where the repo and the data folder live; RPA_DATA_DIR (env or .env) moves the data folder."""
from __future__ import annotations

import os
from pathlib import Path

from worker.naver_client import ENV_FILE, read_env_file

REPO_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = REPO_ROOT / "config.xlsx"
DATA_DIR_KEY = "RPA_DATA_DIR"


def data_dir(env_file: Path = ENV_FILE) -> Path:
    """Env var first, then .env, otherwise the folder that holds the repo."""
    override = os.environ.get(DATA_DIR_KEY) or read_env_file(env_file).get(DATA_DIR_KEY, "")
    return Path(override) if override else REPO_ROOT.parent


DATA_DIR = data_dir()
