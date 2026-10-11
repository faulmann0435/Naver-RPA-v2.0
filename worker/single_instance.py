"""Keep two copies of the same program from running at once (an OS file lock, freed on exit)."""
from __future__ import annotations

import sys
from pathlib import Path
from typing import IO


class AlreadyRunning(RuntimeError):
    """Another process (or handle) already holds the lock."""


class InstanceLock:
    """An exclusive lock on one byte of a file; the OS drops it when the process dies."""

    def __init__(self, path: Path, handle: IO[bytes]):
        self.path = path
        self._handle: IO[bytes] | None = handle

    def release(self) -> None:
        handle, self._handle = self._handle, None
        if handle is None:
            return
        try:
            _unlock(handle)
        finally:
            handle.close()

    def __enter__(self) -> InstanceLock:
        return self

    def __exit__(self, *exc) -> None:
        self.release()


if sys.platform == "win32":
    import msvcrt

    def _lock(handle: IO[bytes]) -> None:
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)

    def _unlock(handle: IO[bytes]) -> None:
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
else:
    import fcntl

    def _lock(handle: IO[bytes]) -> None:
        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)

    def _unlock(handle: IO[bytes]) -> None:
        fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def acquire(name: str, lock_dir: Path) -> InstanceLock:
    """Take the lock `lock_dir/name.lock` or raise AlreadyRunning. Keep the result alive."""
    lock_dir.mkdir(parents=True, exist_ok=True)
    path = lock_dir / f"{name}.lock"
    handle = open(path, "a+b")
    try:
        _lock(handle)
    except OSError as exc:
        handle.close()
        raise AlreadyRunning(f"{name}이(가) 이미 실행 중입니다 ({path})") from exc
    return InstanceLock(path, handle)
