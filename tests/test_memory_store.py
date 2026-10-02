"""MemoryStore implements the DataStore contract."""
import pytest

from store.base import Author, ConflictError
from store.memory_store import MemoryStore


def test_read_missing_is_none():
    assert MemoryStore().read_text("a.csv") is None


def test_create_and_read():
    store = MemoryStore()
    sha = store.write_text("a.csv", "v1", None, "create")
    snap = store.read_text("a.csv")
    assert snap is not None and snap.content == "v1" and snap.sha == sha


def test_update_with_correct_sha():
    store = MemoryStore()
    sha1 = store.write_text("a.csv", "v1", None, "create")
    sha2 = store.write_text("a.csv", "v2", sha1, "update")
    assert sha2 != sha1
    assert store.read_text("a.csv").content == "v2"


def test_stale_sha_conflicts():
    store = MemoryStore()
    sha1 = store.write_text("a.csv", "v1", None, "create")
    store.write_text("a.csv", "v2", sha1, "update")
    with pytest.raises(ConflictError):
        store.write_text("a.csv", "v3", sha1, "stale")


def test_create_when_exists_conflicts():
    store = MemoryStore()
    store.write_text("a.csv", "v1", None, "create")
    with pytest.raises(ConflictError):
        store.write_text("a.csv", "v2", None, "again")


def test_update_missing_file_conflicts():
    with pytest.raises(ConflictError):
        MemoryStore().write_text("a.csv", "v", "abc", "x")


def test_history_newest_first_and_limit():
    store = MemoryStore()
    sha1 = store.write_text("a.csv", "v1", None, "first", Author("kim", "k@x.com"))
    sha2 = store.write_text("a.csv", "v2", sha1, "second")
    store.write_text("a.csv", "v3", sha2, "third")
    hist = store.history("a.csv")
    assert [r.message for r in hist] == ["third", "second", "first"]
    assert hist[-1].author == "kim"
    assert len(store.history("a.csv", limit=2)) == 2
    assert store.history("nope.csv") == []


def test_read_text_at():
    store = MemoryStore()
    sha1 = store.write_text("a.csv", "v1", None, "first")
    store.write_text("a.csv", "v2", sha1, "second")
    assert store.read_text_at("a.csv", sha1) == "v1"
    assert store.read_text_at("a.csv", "unknown") is None


def test_shas_are_deterministic():
    a, b = MemoryStore(), MemoryStore()
    assert a.write_text("f", "x", None, "m") == b.write_text("f", "x", None, "m")
