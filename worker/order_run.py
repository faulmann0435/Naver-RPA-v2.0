"""One 발주 처리 run: recent paid orders from the API -> purchase-order files, like the v15 app.

Uses the same rules, layouts and item dictionary as the app (the data repository named in
.streamlit/secrets.toml), falling back to config.xlsx without a dictionary when there is none.
Orders already marked as shared (the ledger) are left out, so pressing the button twice never
puts an order on two purchase orders. include_shared=True ([📋 발송 안 된 주문 전부]) ignores the
ledger and puts every unshipped order back, like processing the store's Excel export.
"""
from __future__ import annotations

import json
import time
from collections.abc import Callable, Mapping
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
from worker.naver_adapter import (
    KST,
    SHIPPABLE_STATUS,
    details_to_dataframe,
    has_address,
    split_shippable,
)
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
# Naver's status fields of the address-less orders, kept to learn how Naver marks them (no personal data)
ADDRESS_WAITING_FILE = DATA_DIR / "발주기록" / "배송지대기_상태.json"
LOOKBACK_DAYS = 14  # covers a long holiday; the ledger keeps the overlap from being sent twice
ALL_LOOKBACK_DAYS = 30  # [발송 안 된 주문 전부] also reaches pre-orders paid long before shipping
MAX_NAMES = 10  # recipients listed in the summary before "외 N건"
PLACE_ORDER_DONE = "OK"  # placeOrderStatus after 발주확인; the seller center lists those under 발송대기
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
    shared_names: tuple[str, ...] = ()  # recipients of the ledger orders (left out, or put back)
    include_shared: bool = False
    # every unshipped paid order with an address, ledger or not, split like the seller center's screen
    new_orders: int = 0  # 신규주문: 발주확인 not done yet
    waiting_orders: int = 0  # 발송대기: 발주확인 done
    # paid orders with no delivery address yet (선물하기 수락대기 and the like): left out and kept off the
    # ledger, so they come back on their own once the address is in
    address_waiting_names: tuple[str, ...] = ()


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
    """Every product order changed in the last `days` days, whatever the change; split_shippable picks
    the paid ones. Not filtered by lastChangedType=PAYED: that keeps only orders whose LAST change is the
    payment, so a paid order changed afterwards (e.g. its delivery address) was silently missing."""
    # One look-back is a burst of calls (one per day); Naver answers GW.RATE_LIMIT (429) to bursts.
    http = RetryingSession(session, sleep=RETRY_SLEEP)
    now = datetime.now(KST)
    token = fetch_token(load_credentials(), http)
    ids = fetch_changed_ids(token, now - timedelta(days=days), now, None, http)
    return fetch_details(token, ids, http) if ids else []


def _recipient(item: Mapping) -> str:
    """'name(product…)' so the owner can tell which order it is."""
    product_order = item["productOrder"]
    name = (product_order.get("shippingAddress") or {}).get("name") or "이름 없음"
    return f"{name}({str(product_order.get('productName', ''))[:12]})"


def _orderer_label(item: Mapping) -> str:
    """'orderer(product…)': with no address there is no receiver name yet."""
    orderer = (item.get("order") or {}).get("ordererName") or "이름 없음"
    return f"{orderer}({str(item['productOrder'].get('productName', ''))[:12]})"


STATUS_KEY_HINTS = ("status", "type", "gift", "date", "class")  # field names worth keeping in the log


def _status_fields(obj: Mapping | None) -> dict:
    """Scalar fields whose name looks like a status or date; names, phones and addresses never match."""
    return {k: v for k, v in (obj or {}).items()
            if not isinstance(v, (Mapping, list)) and any(h in k.lower() for h in STATUS_KEY_HINTS)}


def log_address_waiting(items: list[Mapping], path: Path = ADDRESS_WAITING_FILE) -> None:
    """Keep Naver's status fields of the address-less orders (latest per order) to learn how
    Naver marks them, e.g. a 선물하기 order the receiver has not accepted yet."""
    if not items:
        return
    logged = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    seen = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    for item in items:
        logged[str(item["productOrder"]["productOrderId"])] = {
            "seen": seen, "order": _status_fields(item.get("order")),
            "productOrder": _status_fields(item.get("productOrder"))}
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(logged, ensure_ascii=False, indent=1), encoding="utf-8")


def run_orders(fetch: Callable[[], list[dict]] = fetch_paid_items,
               engine_loader: Callable[[], Engine] = load_engine,
               ledger: Path = LEDGER_FILE, include_shared: bool = False,
               address_log: Path = ADDRESS_WAITING_FILE) -> RunResult:
    """Fetch, filter, convert and build the purchase orders. Changes nothing on the store."""
    items = fetch()
    shippable, not_shippable = split_shippable(items)
    # paid but held back by a cancel/return/exchange request; shipped or cancelled orders are not news
    claims = [i for i in not_shippable if i["productOrder"].get("productOrderStatus") == SHIPPABLE_STATUS]
    shared = read_ledger(ledger)
    is_shared = lambda i: str(i["productOrder"]["productOrderId"]) in shared
    # no address yet: the store's Excel export and its 신규주문/발송대기 counts leave these out too
    no_address = [i for i in shippable if not has_address(i)]
    ready = [i for i in shippable if has_address(i)]
    log_address_waiting(no_address, address_log)
    in_ledger = [i for i in ready if is_shared(i)]
    confirmed = sum(1 for i in ready if i["productOrder"].get("placeOrderStatus") == PLACE_ORDER_DONE)
    todo = ready if include_shared else [i for i in ready if not is_shared(i)]
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
        already_shared=len(in_ledger),
        unmatched=0 if result is None else len(result.unmatched),
        needs_review=int(stats.get("rows_needs_review", 0)),
        source=engine.source,
        shared_names=tuple(dict.fromkeys(_recipient(i) for i in in_ledger)),
        include_shared=include_shared,
        new_orders=len(ready) - confirmed,
        waiting_orders=confirmed,
        address_waiting_names=tuple(_orderer_label(i) for i in no_address),
    )


def run_all_unshipped() -> RunResult:
    """[발송 안 된 주문 전부]: every unshipped order of the last ALL_LOOKBACK_DAYS, ledger ignored."""
    return run_orders(fetch=lambda: fetch_paid_items(ALL_LOOKBACK_DAYS), include_shared=True)


def _names(names: tuple[str, ...]) -> str:
    shown = ", ".join(names[:MAX_NAMES])
    return shown + (f" 외 {len(names) - MAX_NAMES}건" if len(names) > MAX_NAMES else "")


def _shared_line(run: RunResult) -> str:
    if run.include_shared:
        line = f"· 전에 공유한 주문 {run.already_shared}건도 다시 넣음"
        if run.already_shared:
            line += f": {_names(run.shared_names)}\n  ⚠️ 이미 현장에 보낸 주문이면 두 번 보내지 않게 확인하세요."
        return line
    line = f"· 이미 공유한 주문 {run.already_shared}건 제외"
    return line + (f": {_names(run.shared_names)}" if run.already_shared else "")


def _naver_line(run: RunResult) -> str:
    """The counts to hold against the seller center's screen (신규주문 + 발송대기)."""
    total = run.new_orders + run.waiting_orders
    return f"· 네이버 발송 전 주문 {total}건 = 신규주문 {run.new_orders} + 발송대기 {run.waiting_orders}"


def _address_waiting_lines(run: RunResult) -> list[str]:
    names = run.address_waiting_names
    if not names:
        return []
    return [f"· 🎁 배송지 입력 대기(선물 수락대기 등) {len(names)}건 제외: {_names(names)}",
            "  주소가 들어오면 다음 [발주 처리]에 자동으로 들어갑니다."]


def summary(run: RunResult) -> str:
    """The message that goes with the files."""
    if not run.paid:
        return "\n".join([f"📭 새로 발주할 주문이 없습니다 ({run.when:%m/%d %H:%M})", _naver_line(run),
                          _shared_line(run), *_address_waiting_lines(run)])
    lines = [
        f"📦 발주서 {len(run.files)}개 ({run.when:%m/%d %H:%M})",
        _naver_line(run),
        f"· 이번 발주서에 넣은 주문 {run.paid}건",
        f"· 취소·반품·교환 요청으로 뺀 주문 {run.left_out_claims}건",
        _shared_line(run),
        *_address_waiting_lines(run),
    ]
    if run.unmatched:
        lines.append(f"· ⚠️ 미등록 품목 {run.unmatched}건 → v15 앱에서 등록한 뒤 다시 [발주 처리]를 눌러 주세요.")
    else:
        lines.append("· ✅ 미등록 품목 없음")
    if run.needs_review:
        lines.append(f"· ℹ️ 사전 검토 표시가 남은 품목 {run.needs_review}줄 (발주서에는 정상 반영, 사전 정리용 참고)")
    return "\n".join(lines)
