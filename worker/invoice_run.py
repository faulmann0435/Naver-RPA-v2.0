"""One 송장 run: the courier's tracking file + unshipped orders from the API -> upload files and a dispatch.

prepare_invoices() changes nothing on the store: it matches and builds the 송장 업로드용 xlsx files.
register() is the step that writes tracking numbers to Naver, and only after the owner taps a button.
Two small files under DATA_DIR/발주기록 keep the state:
  송장대기.json       runs waiting for a button tap, so a restart keeps the buttons valid
  송장등록_기록.json   product orders already registered (written for really dispatched ones only)
"""
from __future__ import annotations

import json
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from core.invoice_match import (
    NOTE_NO_ORDER,
    Assignment,
    MatchResult,
    build_upload_workbook,
    match_invoices,
    normalize_orders,
    read_invoice_file,
)
from worker.dispatch_client import DispatchItem, DispatchResult, dispatch_or_simulate
from worker.naver_adapter import KST, details_to_dataframe, split_shippable
from worker.order_run import fetch_paid_items
from worker.paths import DATA_DIR

LEDGER_FILE = DATA_DIR / "발주기록" / "송장등록_기록.json"
PENDING_FILE = DATA_DIR / "발주기록" / "송장대기.json"
MATCHED_FILE_NAME = "송장업로드_일치_{when:%m%d_%H%M}.xlsx"
WITH_CHECKS_FILE_NAME = "송장업로드_확인포함_{when:%m%d_%H%M}.xlsx"
SUMMARY_LIMIT = 3800  # Telegram caps a message at 4096 characters
ID_COLUMN = "상품주문번호"


@dataclass(frozen=True)
class InvoiceFile:
    filename: str
    data: bytes


@dataclass(frozen=True)
class InvoiceRun:
    result: MatchResult
    files: tuple[InvoiceFile, ...]
    already_registered: int = 0
    when: datetime = field(default_factory=lambda: datetime.now(KST))


# ---------------------------------------------------------------- state files

def read_ledger(path: Path = LEDGER_FILE) -> dict[str, dict]:
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


def mark_registered(tracking_of: dict[str, str], when: datetime, path: Path = LEDGER_FILE) -> None:
    ledger = read_ledger(path)
    stamp = f"{when:%Y-%m-%d %H:%M}"
    ledger.update({pid: {"tracking": tracking, "when": stamp} for pid, tracking in tracking_of.items()})
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(ledger, ensure_ascii=False, indent=1), encoding="utf-8")


def read_pending(path: Path = PENDING_FILE) -> dict[str, dict]:
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


def _write_pending(pending: dict[str, dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(pending, ensure_ascii=False, indent=1), encoding="utf-8")


def _groups(assignments: tuple[Assignment, ...]) -> list[list[str]]:
    return [[*a.product_order_ids, a.tracking] for a in assignments]


def run_entry(run: InvoiceRun) -> dict:
    """What a pending run stores: per shipment [product order ids..., tracking number]."""
    return {"when": f"{run.when:%m/%d %H:%M}", "matched": _groups(run.result.matched()),
            "includable": _groups(run.result.includable())}


def add_pending(run: InvoiceRun, path: Path = PENDING_FILE) -> str:
    """Remember a run for its buttons; returns the run id."""
    run_id = uuid.uuid4().hex[:12]
    _write_pending({**read_pending(path), run_id: run_entry(run)}, path)
    return run_id


def remove_pending(run_id: str, path: Path = PENDING_FILE) -> None:
    _write_pending({k: v for k, v in read_pending(path).items() if k != run_id}, path)


# ---------------------------------------------------------------- run

def prepare_invoices(invoice_bytes: bytes, fetch: Callable[[], list[dict]] = fetch_paid_items,
                     ledger: Path = LEDGER_FILE, when: datetime | None = None) -> InvoiceRun:
    """Match the courier file against the unshipped orders and build the upload files."""
    invoices = read_invoice_file(invoice_bytes)
    shippable, _claims = split_shippable(fetch())
    table = normalize_orders(details_to_dataframe(shippable))
    done = set(read_ledger(ledger))
    fresh = table[~table[ID_COLUMN].isin(done)]
    result = match_invoices(fresh, invoices)
    moment = when or datetime.now(KST)
    files = []
    if result.matched():
        files.append(InvoiceFile(MATCHED_FILE_NAME.format(when=moment),
                                 build_upload_workbook(fresh, result.matched())))
    if result.includable_checks():
        files.append(InvoiceFile(WITH_CHECKS_FILE_NAME.format(when=moment),
                                 build_upload_workbook(fresh, result.includable())))
    return InvoiceRun(result, tuple(files), len(table) - len(fresh), moment)


def register(entry: dict, include_checks: bool,
             dispatch: Callable[[list[DispatchItem]], DispatchResult] = dispatch_or_simulate,
             ledger: Path = LEDGER_FILE, now: Callable[[], datetime] = lambda: datetime.now(KST)) -> DispatchResult:
    """Send the pending run's tracking numbers to Naver; record the really registered ones."""
    done = set(read_ledger(ledger))
    groups = entry["includable" if include_checks else "matched"]
    tracking_of = {pid: group[-1] for group in groups for pid in group[:-1] if pid not in done}
    result = dispatch([DispatchItem(pid, tracking) for pid, tracking in tracking_of.items()])
    if not result.simulated and result.success_ids:
        mark_registered({pid: tracking_of[pid] for pid in result.success_ids if pid in tracking_of},
                        now(), ledger)
    return result


# ---------------------------------------------------------------- message

def _clip(lines: list[str], budget: int) -> list[str]:
    kept: list[str] = []
    used = 0
    for index, line in enumerate(lines):
        if used + len(line) + 1 > budget:
            return [*kept, f"… 그 외 {len(lines) - index}건"]
        kept.append(line)
        used += len(line) + 1
    return kept


def invoice_summary(run: InvoiceRun) -> str:
    """Korean message that goes with the files; kept under Telegram's size limit."""
    result = run.result
    checks, loose = result.needs_check(), result.invoices_without_order
    lines = [
        f"🚚 송장 확인 ({run.when:%m/%d %H:%M})",
        f"· 일치 {len(result.matched())}건",
        f"· 확인 필요 {len(checks)}건 (그중 포함 가능 {len(result.includable_checks())}건)",
        f"· 주문 못 찾은 송장 {len(loose)}건",
        f"· 송장 없는 발송대기 주문 {len(result.orders_without_invoice)}건",
    ]
    if result.invoices_without_tracking:
        lines.append(f"· ⚠️ 송장번호가 없는 줄 {len(result.invoices_without_tracking)}개 (보낸 곳에 확인)")
    if run.already_registered:
        lines.append(f"· 이미 등록한 주문 {run.already_registered}건 제외")
    detail: list[str] = []
    if checks:
        detail += ["", "⚠️ 확인 필요"] + [
            f"· {a.recipient} · {a.reason} · {a.tracking}" + ("" if a.includable else " (제외)") for a in checks]
    if result.invoices_without_tracking:
        detail += ["", "⚠️ 송장번호가 없는 줄"] + [
            f"· {i.seq}번 줄 · {i.name or '받는 사람 없음'} · {i.item_name}" for i in result.invoices_without_tracking]
    if loose:
        detail += ["", f"❓ {NOTE_NO_ORDER}"] + [f"· {i.seq} · {i.name} · {i.tracking}" for i in loose]
    if result.orders_without_invoice:
        detail += ["", "📭 송장 없는 주문"] + [f"· {o.recipient} · {o.items}" for o in result.orders_without_invoice]
    head = "\n".join(lines)
    return head + "\n" + "\n".join(_clip(detail, SUMMARY_LIMIT - len(head)))
