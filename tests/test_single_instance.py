"""The single-instance lock: a second handle is refused until the first is released."""
from __future__ import annotations

import pytest

from worker.single_instance import AlreadyRunning, acquire


def test_second_acquire_fails(tmp_path):
    first = acquire("bot", tmp_path)
    try:
        with pytest.raises(AlreadyRunning):
            acquire("bot", tmp_path)
    finally:
        first.release()


def test_acquire_again_after_release(tmp_path):
    acquire("bot", tmp_path).release()
    with acquire("bot", tmp_path):
        pass


def test_different_names_do_not_clash(tmp_path):
    with acquire("a", tmp_path), acquire("b", tmp_path):
        pass


def test_creates_missing_dir_and_release_is_idempotent(tmp_path):
    lock = acquire("bot", tmp_path / "x" / "y")
    lock.release()
    lock.release()
    assert (tmp_path / "x" / "y" / "bot.lock").exists()
