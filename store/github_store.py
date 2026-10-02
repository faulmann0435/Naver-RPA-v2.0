"""DataStore backed by the GitHub REST contents API."""
from __future__ import annotations

import base64
from collections.abc import Mapping
from typing import Any
from urllib.parse import quote

import requests

from store.base import Author, ConflictError, Revision, Snapshot, StoreError

API_ROOT = "https://api.github.com"


class GitHubStore:
    """Reads and writes text files in a GitHub repository branch."""

    def __init__(
        self,
        repo: str,
        branch: str,
        token: str,
        session: requests.Session | None = None,
        timeout: float = 15.0,
    ) -> None:
        self.repo = repo
        self.branch = branch
        self._token = token
        self._session: Any = session if session is not None else requests.Session()
        self._timeout = timeout

    def __repr__(self) -> str:
        return f"GitHubStore(repo={self.repo!r}, branch={self.branch!r}, token='***')"

    @classmethod
    def from_secrets(cls, section: Mapping[str, str]) -> GitHubStore:
        """Build from a secrets section with keys repo, branch, token."""
        try:
            return cls(repo=section["repo"], branch=section["branch"], token=section["token"])
        except KeyError as e:
            raise StoreError(f"데이터 저장소 설정에 {e.args[0]!r} 항목이 없습니다.") from None

    # -- helpers -----------------------------------------------------------
    @property
    def _headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self._token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        }

    def _contents_url(self, path: str) -> str:
        return f"{API_ROOT}/repos/{self.repo}/contents/{quote(path)}"

    @staticmethod
    def _message(resp: Any) -> str:
        try:
            data = resp.json()
            msg = data.get("message", "") if isinstance(data, dict) else ""
        except ValueError:
            msg = ""
        return str(msg or getattr(resp, "text", ""))[:300]

    def _fail(self, resp: Any, action: str) -> StoreError:
        text = f"GitHub {action} 실패 (HTTP {resp.status_code}): {self._message(resp)}"
        return StoreError(text.replace(self._token, "***") if self._token else text)

    def _get(self, url: str, params: dict[str, Any] | None = None) -> Any:
        try:
            return self._session.get(url, headers=self._headers, params=params, timeout=self._timeout)
        except requests.RequestException as e:
            raise StoreError(f"GitHub 연결 실패: {type(e).__name__}") from None

    @staticmethod
    def _decode(b64: str) -> str:
        # utf-8 (not utf-8-sig): a leading BOM is part of the file and must be kept
        return base64.b64decode(b64).decode("utf-8")

    def _fetch_contents(self, path: str, ref: str) -> tuple[str, str] | None:
        """Return (text, sha) at ref, or None if missing."""
        resp = self._get(self._contents_url(path), {"ref": ref})
        if resp.status_code == 404:
            return None
        if not resp.ok:
            raise self._fail(resp, "읽기")
        data = resp.json()
        sha = str(data["sha"])
        encoded = data.get("content") or ""
        if not encoded:  # files over 1 MB come back without content
            blob = self._get(f"{API_ROOT}/repos/{self.repo}/git/blobs/{sha}")
            if not blob.ok:
                raise self._fail(blob, "대용량 읽기")
            encoded = blob.json().get("content", "")
        return self._decode(encoded), sha

    # -- DataStore ---------------------------------------------------------
    def read_text(self, path: str) -> Snapshot | None:
        found = self._fetch_contents(path, self.branch)
        return None if found is None else Snapshot(content=found[0], sha=found[1])

    def write_text(
        self,
        path: str,
        content: str,
        expected_sha: str | None,
        message: str,
        author: Author | None = None,
    ) -> str:
        body: dict[str, Any] = {
            "branch": self.branch,
            "message": message,
            "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
        }
        if expected_sha is not None:
            body["sha"] = expected_sha
        if author is not None:
            person = {"name": author.name, "email": author.email}
            body["author"] = person
            body["committer"] = person
        try:
            resp = self._session.put(
                self._contents_url(path), headers=self._headers, json=body, timeout=self._timeout
            )
        except requests.RequestException as e:
            raise StoreError(f"GitHub 연결 실패: {type(e).__name__}") from None
        is_sha_422 = resp.status_code == 422 and "sha" in self._message(resp).lower()
        if resp.status_code == 409 or is_sha_422:
            raise ConflictError(f"{path}: 다른 사람이 먼저 저장했거나 이미 존재합니다. ({self._message(resp)})")
        if not resp.ok:
            raise self._fail(resp, "저장")
        return str(resp.json()["content"]["sha"])

    def history(self, path: str, limit: int = 30) -> list[Revision]:
        resp = self._get(
            f"{API_ROOT}/repos/{self.repo}/commits",
            {"path": path, "sha": self.branch, "per_page": limit},
        )
        if not resp.ok:
            raise self._fail(resp, "이력 조회")
        revisions: list[Revision] = []
        for item in resp.json():
            commit = item.get("commit", {})
            who = commit.get("author") or {}
            revisions.append(
                Revision(
                    sha=str(item.get("sha", "")),
                    date=str(who.get("date", "")),
                    author=str(who.get("name", "")),
                    message=str(commit.get("message", "")),
                )
            )
        return revisions

    def read_text_at(self, path: str, ref: str) -> str | None:
        found = self._fetch_contents(path, ref)
        return None if found is None else found[0]
