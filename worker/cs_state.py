"""Files of the CS (customer question) feature that live in the CS state dir, plus its on/off switch.

  notified.json     {key: entry} one entry per question seen (see worker.cs_watch)
  pairs.jsonl       (AI draft, owner answer) for questions answered in the seller center
  outcomes.jsonl    one line per settled question (see worker.cs_outcomes)
  cs_pending.json   what the Telegram flow waits for: {"waiting": {chat: {sid, since}},
                    "previews": {"chat:sid": {text, source, since}}}; entries expire after 24 hours

Kept apart from worker.cs_watch so the Telegram flow can use it without importing the watcher.
"""
from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from worker.naver_client import ENV_FILE, read_env_file

WATCH_SWITCH_KEY = "CS_WATCH_ENABLED"
NOTIFIED_FILE, PAIRS_FILE = "notified.json", "pairs.jsonl"
OUTCOMES_FILE, PENDING_FILE = "outcomes.jsonl", "cs_pending.json"
CHANNEL_QNA, CHANNEL_INQUIRY = "상품문의", "고객문의"
PENDING_TTL = timedelta(hours=24)
SID_LENGTH = 8


def cs_enabled() -> bool:
    """True when CS_WATCH_ENABLED=1 is set in the environment or .env (read on every call)."""
    value = os.environ.get(WATCH_SWITCH_KEY) or read_env_file(ENV_FILE).get(WATCH_SWITCH_KEY, "")
    return value.strip() == "1"


def short_id(key: str) -> str:
    """Short stable id of a state key; fits Telegram's 64-byte callback data."""
    return hashlib.sha1(key.encode()).hexdigest()[:SID_LENGTH]


# --- notified.json / pairs.jsonl / outcomes.jsonl ---------------------------------------------

def load_state(state_dir: Path) -> tuple[dict[str, dict], bool]:
    """(state, first_run): first_run is True when notified.json does not exist yet."""
    path = state_dir / NOTIFIED_FILE
    if not path.exists():
        return {}, True
    return json.loads(path.read_text(encoding="utf-8")), False


def save_state(state_dir: Path, state: dict[str, dict]) -> None:
    state_dir.mkdir(parents=True, exist_ok=True)
    tmp = state_dir / (NOTIFIED_FILE + ".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(state_dir / NOTIFIED_FILE)


def append_jsonl(state_dir: Path, name: str, record: dict) -> None:
    state_dir.mkdir(parents=True, exist_ok=True)
    with (state_dir / name).open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False) + "\n")


def append_pair(state_dir: Path, pair: dict) -> None:
    append_jsonl(state_dir, PAIRS_FILE, pair)


def read_jsonl(state_dir: Path, name: str) -> list[dict]:
    path = state_dir / name
    if not path.exists():
        return []
    lines = (line for line in path.read_text(encoding="utf-8").splitlines() if line.strip())
    return [json.loads(line) for line in lines]


def find_by_sid(state: dict[str, dict], sid: str) -> tuple[str, dict] | None:
    return next(((k, e) for k, e in state.items() if e.get("sid") == sid), None)


def find_by_message(state: dict[str, dict], chat: int, message_id: int) -> tuple[str, dict] | None:
    """The question whose alert is the given Telegram message of the given chat."""
    return next(((k, e) for k, e in state.items()
                 if (e.get("messages") or {}).get(str(chat)) == message_id), None)


# --- cs_pending.json --------------------------------------------------------------------------

def _fresh(entries: dict[str, dict], now: datetime) -> dict[str, dict]:
    def alive(entry: dict) -> bool:
        try:
            return now - datetime.fromisoformat(entry["since"]) < PENDING_TTL
        except (KeyError, ValueError, TypeError):
            return False
    return {k: v for k, v in entries.items() if alive(v)}


def read_pending(state_dir: Path, now: datetime) -> dict[str, dict[str, Any]]:
    """{"waiting": ..., "previews": ...} without the expired entries (a missing or broken file is empty)."""
    path = state_dir / PENDING_FILE
    try:
        data = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    except ValueError:
        data = {}
    return {part: _fresh(dict(data.get(part) or {}), now) for part in ("waiting", "previews")}


def write_pending(state_dir: Path, pending: dict[str, dict]) -> None:
    state_dir.mkdir(parents=True, exist_ok=True)
    tmp = state_dir / (PENDING_FILE + ".tmp")
    tmp.write_text(json.dumps(pending, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(state_dir / PENDING_FILE)
