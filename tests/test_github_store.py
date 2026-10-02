"""GitHubStore against a fake requests session (offline)."""
import base64
from typing import Any

import pytest
import requests

from store.base import Author, ConflictError, StoreError
from store.github_store import GitHubStore

TOKEN = "ghp_SECRET_TOKEN_123"


class FakeResponse:
    def __init__(self, status_code: int, payload: Any = None, text: str = ""):
        self.status_code = status_code
        self._payload = payload
        self.text = text

    @property
    def ok(self) -> bool:
        return 200 <= self.status_code < 300

    def json(self) -> Any:
        if self._payload is None:
            raise ValueError("no json")
        return self._payload


class FakeSession:
    def __init__(self, responses: list[FakeResponse] | None = None, error: Exception | None = None):
        self.responses = list(responses or [])
        self.error = error
        self.calls: list[tuple[str, str, dict[str, Any]]] = []

    def _next(self, method: str, url: str, kwargs: dict[str, Any]) -> FakeResponse:
        self.calls.append((method, url, kwargs))
        if self.error:
            raise self.error
        return self.responses.pop(0)

    def get(self, url: str, **kwargs: Any) -> FakeResponse:
        return self._next("GET", url, kwargs)

    def put(self, url: str, **kwargs: Any) -> FakeResponse:
        return self._next("PUT", url, kwargs)


def _b64(text: str) -> str:
    return base64.b64encode(text.encode("utf-8")).decode("ascii")


def _store(*responses: FakeResponse, error: Exception | None = None) -> tuple[GitHubStore, FakeSession]:
    session = FakeSession(list(responses), error)
    return GitHubStore("owner/repo", "dev", TOKEN, session=session), session  # type: ignore[arg-type]


def test_read_200_decodes_and_keeps_bom():
    store, session = _store(FakeResponse(200, {"content": _b64("﻿한글,a\n1,2\n"), "sha": "abc"}))
    snap = store.read_text("product_route.csv")
    assert snap is not None
    assert snap.content == "﻿한글,a\n1,2\n" and snap.sha == "abc"
    _, url, kwargs = session.calls[0]
    assert url == "https://api.github.com/repos/owner/repo/contents/product_route.csv"
    assert kwargs["params"] == {"ref": "dev"}
    assert kwargs["headers"]["Authorization"] == f"Bearer {TOKEN}"
    assert kwargs["headers"]["Accept"] == "application/vnd.github+json"
    assert kwargs["headers"]["X-GitHub-Api-Version"] == "2022-11-28"


def test_read_404_is_none():
    store, _ = _store(FakeResponse(404, {"message": "Not Found"}))
    assert store.read_text("x.csv") is None


def test_large_file_blob_fallback():
    store, session = _store(
        FakeResponse(200, {"content": "", "sha": "big"}),
        FakeResponse(200, {"content": _b64("large body")}),
    )
    snap = store.read_text("big.csv")
    assert snap is not None and snap.content == "large body" and snap.sha == "big"
    assert session.calls[1][1].endswith("/repos/owner/repo/git/blobs/big")


def test_put_body_has_sha_branch_author():
    store, session = _store(FakeResponse(200, {"content": {"sha": "new"}}))
    sha = store.write_text("a.csv", "내용", "old", "msg", Author("kim", "k@x.com"))
    assert sha == "new"
    method, _, kwargs = session.calls[0]
    body = kwargs["json"]
    assert method == "PUT"
    assert body["sha"] == "old" and body["branch"] == "dev" and body["message"] == "msg"
    assert base64.b64decode(body["content"]).decode("utf-8") == "내용"
    assert body["author"] == {"name": "kim", "email": "k@x.com"}
    assert body["committer"] == {"name": "kim", "email": "k@x.com"}


def test_create_omits_sha_and_author():
    store, session = _store(FakeResponse(201, {"content": {"sha": "n"}}))
    store.write_text("a.csv", "x", None, "create")
    body = session.calls[0][2]["json"]
    assert "sha" not in body and "author" not in body


def test_409_is_conflict():
    store, _ = _store(FakeResponse(409, {"message": "conflict"}))
    with pytest.raises(ConflictError):
        store.write_text("a.csv", "x", "old", "m")


def test_422_sha_message_is_conflict():
    store, _ = _store(FakeResponse(422, {"message": "Invalid request. \"sha\" wasn't supplied."}))
    with pytest.raises(ConflictError):
        store.write_text("a.csv", "x", None, "m")


def test_422_other_is_store_error_not_conflict():
    store, _ = _store(FakeResponse(422, {"message": "bad content"}))
    with pytest.raises(StoreError) as exc:
        store.write_text("a.csv", "x", None, "m")
    assert not isinstance(exc.value, ConflictError)


def test_500_is_store_error_with_status():
    store, _ = _store(FakeResponse(500, {"message": "boom"}))
    with pytest.raises(StoreError) as exc:
        store.write_text("a.csv", "x", "old", "m")
    assert "500" in str(exc.value) and "boom" in str(exc.value)


def test_token_not_leaked():
    store, _ = _store(FakeResponse(401, {"message": f"Bad credentials {TOKEN}"}))
    with pytest.raises(StoreError) as exc:
        store.read_text("a.csv")
    assert TOKEN not in str(exc.value)
    assert TOKEN not in repr(store) and TOKEN not in str(store)


def test_network_error_wrapped_without_token():
    store, _ = _store(error=requests.ConnectionError(f"failed {TOKEN}"))
    with pytest.raises(StoreError) as exc:
        store.read_text("a.csv")
    assert TOKEN not in str(exc.value)


def test_history_parsed():
    items = [
        {"sha": "c2", "commit": {"author": {"name": "kim", "date": "2026-01-02T00:00:00Z"}, "message": "second"}},
        {"sha": "c1", "commit": {"author": {"name": "lee", "date": "2026-01-01T00:00:00Z"}, "message": "first"}},
    ]
    store, session = _store(FakeResponse(200, items))
    revs = store.history("a.csv", limit=5)
    assert [r.sha for r in revs] == ["c2", "c1"] and revs[0].author == "kim"
    params = session.calls[0][2]["params"]
    assert params == {"path": "a.csv", "sha": "dev", "per_page": 5}


def test_read_text_at_uses_commit_ref():
    store, session = _store(FakeResponse(200, {"content": _b64("old"), "sha": "s"}))
    assert store.read_text_at("a.csv", "c1") == "old"
    assert session.calls[0][2]["params"] == {"ref": "c1"}


def test_from_secrets():
    store = GitHubStore.from_secrets({"repo": "o/r", "branch": "main", "token": TOKEN})
    assert store.repo == "o/r" and store.branch == "main"
    with pytest.raises(StoreError):
        GitHubStore.from_secrets({"repo": "o/r"})
