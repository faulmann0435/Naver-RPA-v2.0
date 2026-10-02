"""Store interface: snapshots, revisions, errors and the DataStore protocol."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol


@dataclass(frozen=True)
class Snapshot:
    """File content together with its version fingerprint."""

    content: str
    sha: str


@dataclass(frozen=True)
class Revision:
    """One entry of a file's change history."""

    sha: str
    date: str
    author: str
    message: str


@dataclass(frozen=True)
class Author:
    """Commit author."""

    name: str
    email: str


class StoreError(Exception):
    """Any storage failure."""


class ConflictError(StoreError):
    """The file changed since it was read (or already exists on create)."""


class DataStore(Protocol):
    """Text-file storage with versioning and optimistic concurrency."""

    def read_text(self, path: str) -> Snapshot | None:
        """Return the file, or None if it does not exist."""
        ...

    def write_text(
        self,
        path: str,
        content: str,
        expected_sha: str | None,
        message: str,
        author: Author | None = None,
    ) -> str:
        """Write the file and return the new sha.

        expected_sha=None creates a new file (ConflictError if it exists);
        a mismatching sha raises ConflictError.
        """
        ...

    def history(self, path: str, limit: int = 30) -> list[Revision]:
        """Return revisions of the file, newest first."""
        ...

    def read_text_at(self, path: str, ref: str) -> str | None:
        """Return the file content at the given commit, or None."""
        ...
