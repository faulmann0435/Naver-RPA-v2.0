"""One 발주 처리 run: recent paid orders from the API -> purchase-order files, like the v15 app.

Uses the same rules, layouts and item dictionary as the app (the data repository named in
.streamlit/secrets.toml), falling back to config.xlsx without a dictionary when there is none.
Orders already marked as shared (the ledger) are left out, so pressing the button twice never
puts an order on two purchase orders.
"""
from __future__ import annotations

import json
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

import requests
import tomllib

from core.config_loader import load_config
from core.dictionary import DictionarySettings, ItemDictionary
from core.pipeline import ProcessResult, process_orders
from store.dictionary_repo import load_state
from store.github_store import GitHubStore
from store.rules_repo import load_config_from_store
from worker.naver_adapter import KST, details_to_dataframe, split_shippable
from worker.naver_client import (
    fetch_changed_ids,
    fetch_details,
    fetch_token,
    load_credentials,
)
from worker.order_link import RetryingSession
from worker.paths import CONFIG_PATH, DATA_DIR, REPO_ROOT

CONFIG_PASSWORD = "1111"  # protects config.xlsx sheets against accidental edits, not a secret
SECRETS_FILE = REPO_ROOT / ".streamlit" / "secrets.toml"
LEDGER_FILE = DATA_DIR / "발주기록" / "공유완료_주문.json"
PENDING_FILE = DATA_DIR / "발주기록" / "공유대기_발주.json"
LOOKBACK_DAYS = 14  # covers a long holiday; the ledger keeps the overlap from being sent twice
RETRY_SLEEP = time.sleep  # wait used between retries after Naver's rate-limit refusal (swapped in tests)


@dataclass(frozen=True)
class Engine:
    config: dict
    dictionary: ItemDictionary
    settings: DictionarySettings
    source: str


@dataclass(frozen=True)
class RunResult:
    files: list[dict]
    product_order_ids: tuple[str, ...]
    paid: int
    left_out_claims: int
    already_shared: int
    unmatched: int
    needs_review: int
    source: str
    when: datetime = field(default_factory=lambda: datetime.now(KST))


def load_engine(secrets_file: Path = SECRETS_FILE) -> Engine:
    """Rules, layouts and dictionary from the app's data repository; config.xlsx when none is set."""
    if secrets_file.exists():
        section = tomllib.loads(secrets_file.read_text(encoding="utf-8")).get("data_store")
        if section:
            store = GitHubStore.from_secrets(section)
            dictionary, settings, _ = load_state(store)
            config = load_config_from_store(store, str(CONFIG_PATH), password=CONFIG_PASSWORD)
            return Engine(config, dictionary, settings, f"데이터 저장소 {store.repo}@{store.branch}")
    return Engine(load_config(str(CONFIG_PATH), password=CONFIG_PASSWORD), ItemDictionary.empty(),
                  DictionarySettings(), "config.xlsx (사전 없음)")


def read_ledger(path: Path = LEDGER_FILE) -> set[str]:
    return set(json.loads(path.read_text(encoding="utf-8"))) if path.exists() else set()


def mark_shared(ids: tuple[str, ...], path: Path = LEDGER_FILE) -> int:
    """Add product order ids to the ledger; returns how many were new."""
    before = read_ledger(path)
    after = before | set(ids)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(sorted(after), ensure_ascii=False, indent=1), encoding="utf-8")
    return len(after) - len(before)


def read_pending(path: Path = PENDING_FILE) -> dict[str, dict]:
    """Runs sent to the phone but not yet confirmed: run id -> {"ids": [...], "when": "MM/DD HH:MM"}."""
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


def write_pending(pending: dict[str, dict], path: Path = PENDING_FILE) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(pending, ensure_ascii=False, indent=1), encoding="utf-8")


def fetch_paid_items(days: int = LOOKBACK_DAYS, session: requests.Session | None = None) -> list[dict]:
    # One look-back is a burst of calls (one per day); Naver answers GW.RATE_LIMIT (429) to bursts.
    http = RetryingSession(session, sleep=RETRY_SLEEP)
    now = datetime.now(KST)
    token = fetch_token(load_credentials(), http)
    return fetch_details(token, fetch_changed_ids(token, now - timedelta(days=days), now, session=http), http)


def run_orders(fetch: Callable[[], list[dict]] = fetch_paid_items,
               engine_loader: Callable[[], Engine] = load_engine,
               ledger: Path = LEDGER_FILE) -> RunResult:
    """Fetch, filter, convert and build the purchase orders. Changes nothing on the store."""
    items = fetch()
    shippable, claims = split_shippable(items)
    shared = read_ledger(ledger)
    todo = [i for i in shippable if str(i["productOrder"]["productOrderId"]) not in shared]
    engine = engine_loader()
    result: ProcessResult | None = None
    if todo:
        result = process_orders(details_to_dataframe(todo), engine.config, engine.dictionary, engine.settings)
    stats = result.stats if result else {}
    return RunResult(
        files=result.files if result else [],
        product_order_ids=tuple(str(i["productOrder"]["productOrderId"]) for i in todo),
        paid=len(todo),
        left_out_claims=len(claims),
        already_shared=len(shippable) - len(todo),
        unmatched=0 if result is None else len(result.unmatched),
        needs_review=int(stats.get("rows_needs_review", 0)),
        source=engine.source,
    )


def summary(run: RunResult) -> str:
    """The message that goes with the files."""
    if not run.paid:
        return f"📭 새로 발주할 주문이 없습니다 ({run.when:%m/%d %H:%M})\n· 이미 공유한 주문 {run.already_shared}건"
    warn = run.unmatched or run.needs_review
    lines = [
        f"📦 발주서 {len(run.files)}개 ({run.when:%m/%d %H:%M})",
        f"· 결제완료 {run.paid}건 반영",
        f"· 취소·반품·교환 요청으로 뺀 주문 {run.left_out_claims}건",
        f"· 이미 공유한 주문 {run.already_shared}건 제외",
        f"· {'⚠️' if run.unmatched else '✅'} 미등록 품목 {run.unmatched}건"
        + (f" · 확인 필요 {run.needs_review}줄" if run.needs_review else ""),
    ]
    if warn:
        lines.append("미등록 품목은 v15 앱에서 등록한 뒤 다시 [발주 처리]를 눌러 주세요.")
    return "\n".join(lines)
