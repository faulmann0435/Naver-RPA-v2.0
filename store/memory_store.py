"""In-memory DataStore used by tests."""
from __future__ import annotations

import hashlib

from store.base import Author, ConflictError, Revision, Snapshot


class MemoryStore:
    """Fully in-memory DataStore with deterministic shas and recorded history."""

    def __init__(self) -> None:
        self._files: dict[str, Snapshot] = {}
        self._revisions: dict[str, list[tuple[Revision, str]]] = {}
        self._counter = 0

    def _next_sha(self, content: str) -> str:
        self._counter += 1
        return hashlib.sha1(f"{self._counter}:{content}".encode()).hexdigest()

    def read_text(self, path: str) -> Snapshot | None:
        return self._files.get(path)

    def write_text(
        self,
        path: str,
        content: str,
        expected_sha: str | None,
        message: str,
        author: Author | None = None,
    ) -> str:
        current = self._files.get(path)
        if expected_sha is None and current is not None:
            raise ConflictError(f"{path} already exists")
        if expected_sha is not None and (current is None or current.sha != expected_sha):
            raise ConflictError(f"{path} changed since it was read")
        sha = self._next_sha(content)
        self._files[path] = Snapshot(content=content, sha=sha)
        rev = Revision(
            sha=sha,
            date=f"2000-01-01T00:00:{self._counter % 60:02d}Z",
            author=author.name if author else "unknown",
            message=message,
        )
        self._revisions.setdefault(path, []).append((rev, content))
        return sha

    def history(self, path: str, limit: int = 30) -> list[Revision]:
        entries = self._revisions.get(path, [])
        return [rev for rev, _ in reversed(entries)][:limit]

    def read_text_at(self, path: str, ref: str) -> str | None:
        for rev, content in self._revisions.get(path, []):
            if rev.sha == ref:
                return content
        return None
