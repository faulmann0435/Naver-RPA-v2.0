"""One 송장 run: the courier's tracking file + unshipped orders from the API -> upload files and a dispatch.

prepare_invoices() changes nothing on the store: it matches and builds the 송장 업로드용 xlsx files.
register() is the step that writes tracking numbers to Naver, and only after the owner taps a button.
Two small files under DATA_DIR/발주기록 keep the state:
  송장대기.json       runs waiting for a button tap, so a restart keeps the buttons valid
  송장등록_기록.json   product orders already registered (written for really dispatched ones only)
Orders of the second channel 자연이랑 (saved by worker.withnature_run) are matched in the same pass, but
they never reach Naver: their tracking numbers go to a 자연이랑_송장업로드 xlsx for the supplier portal, and
the (주문번호, 운송장번호) pairs put in such a file are kept in 자연이랑_송장기록.json.
"""
from __future__ import annotations

import json
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path

import pandas as pd

from core.invoice_match import (
    COURIER_CODES,
    NOTE_NO_ORDER,
    STATUS_CHECK,
    Assignment,
    MatchResult,
    UnshippedOrder,
    build_upload_workbook,
    match_invoices,
    normalize_orders,
    read_invoice_file,
)
from core.withnature import CHANNEL, build_portal_upload, courier_code
from worker.dispatch_client import DispatchItem, DispatchResult, dispatch_or_simulate
from worker.naver_adapter import KST, details_to_dataframe, split_shippable
from worker.order_run import fetch_paid_items
from worker.paths import DATA_DIR
from worker.withnature_run import load_saved_orders

LEDGER_FILE = DATA_DIR / "발주기록" / "송장등록_기록.json"
PENDING_FILE = DATA_DIR / "발주기록" / "송장대기.json"
MATCHED_FILE_NAME = "송장업로드_일치_{when:%m%d_%H%M}.xlsx"
WITH_CHECKS_FILE_NAME = "송장업로드_확인포함_{when:%m%d_%H%M}.xlsx"
WITHNATURE_LEDGER_FILE = DATA_DIR / "발주기록" / "자연이랑_송장기록.json"
WITHNATURE_FILE_NAME = f"{CHANNEL}_송장업로드_{{when:%m%d_%H%M}}.xlsx"
WITHNATURE_CHECKS_FILE_NAME = f"{CHANNEL}_송장업로드_확인필요_{{when:%m%d_%H%M}}.xlsx"
CHANNEL_COLUMN = "_channel"
NAVER_CHANNEL = "naver"
MAX_LISTED = 10
SUMMARY_LIMIT = 3800  # Telegram caps a message at 4096 characters
ID_COLUMN = "상품주문번호"
MAX_FAILURE_LINES = 20


@dataclass(frozen=True)
class InvoiceFile:
    filename: str
    data: bytes


@dataclass(frozen=True)
class WithnatureShipments:
    """What the tracking file did to the saved 자연이랑 orders (nothing goes to Naver)."""
    matched: int  # rows in the 일치 upload file
    checks: int  # rows in the 확인필요 upload file (includable ones only)
    skipped: int  # pairs already put in an earlier upload file
    unknown_courier: tuple[str, ...]  # "recipient · tracking" left out: the courier is not in the portal's list
    excluded: tuple[str, ...]  # "recipient · reason" of needs-check shipments that cannot be included
    unshipped: int  # saved orders without a tracking number in this file


@dataclass(frozen=True)
class InvoiceRun:
    result: MatchResult
    files: tuple[InvoiceFile, ...]
    already_registered: int = 0
    when: datetime = field(default_factory=lambda: datetime.now(KST))
    withnature: WithnatureShipments | None = None


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


def _groups(assignments: tuple[Assignment, ...]) -> list[dict]:
    return [{"ids": list(a.product_order_ids), "tracking": a.tracking, "courier": COURIER_CODES[a.carrier],
             "recipient": a.recipient} for a in assignments]


def _read_group(group) -> tuple[list[str], str, str]:
    """(ids, tracking, courier code); older pending files stored [ids..., tracking] for 한진 only."""
    if isinstance(group, dict):
        return list(group["ids"]), group["tracking"], group["courier"]
    return list(group[:-1]), group[-1], COURIER_CODES["한진택배"]


def entry_items(entry: dict, include_checks: bool) -> list[dict]:
    """The shipments register() would send for this pending entry, shaped like dispatch_queue items."""
    groups = entry["includable" if include_checks else "matched"]
    return [{"ids": list(ids), "tracking": tracking, "courier": courier,
             "recipient": str(group.get("recipient", "")) if isinstance(group, dict) else ""}
            for group in groups for ids, tracking, courier in [_read_group(group)]]


def run_entry(run: InvoiceRun) -> dict:
    """What a pending run stores: per shipment its product order ids, tracking number and courier code."""
    return {"when": f"{run.when:%m/%d %H:%M}", "matched": _groups(run.result.matched()),
            "includable": _groups(run.result.includable())}


def add_pending(run: InvoiceRun, path: Path = PENDING_FILE) -> str:
    """Remember a run for its buttons; returns the run id."""
    run_id = uuid.uuid4().hex[:12]
    _write_pending({**read_pending(path), run_id: run_entry(run)}, path)
    return run_id


def remove_pending(run_id: str, path: Path = PENDING_FILE) -> None:
    _write_pending({k: v for k, v in read_pending(path).items() if k != run_id}, path)


# ---------------------------------------------------------------- 자연이랑

def read_withnature_ledger(path: Path | None = None) -> dict[str, str]:
    """'주문번호|운송장번호' -> when it was put in an upload file."""
    path = path or WITHNATURE_LEDGER_FILE
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


def mark_withnature_uploaded(pairs: list[tuple[str, str]], when: datetime, path: Path | None = None) -> None:
    path = path or WITHNATURE_LEDGER_FILE
    ledger = read_withnature_ledger(path)
    stamp = f"{when:%Y-%m-%d %H:%M}"
    ledger.update({f"{order}|{tracking}": stamp for order, tracking in pairs})
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(ledger, ensure_ascii=False, indent=1), encoding="utf-8")


def _with_channels(naver: pd.DataFrame, withnature: pd.DataFrame) -> pd.DataFrame:
    """Both order tables in one, each row marked with its channel."""
    parts = [frame.assign(**{CHANNEL_COLUMN: name})
             for frame, name in ((naver, NAVER_CHANNEL), (withnature, CHANNEL)) if not frame.empty]
    return normalize_orders(pd.concat(parts, ignore_index=True))


def _flag_cross_channel(assignments: tuple[Assignment, ...], channel_of: dict[str, str]) -> tuple[Assignment, ...]:
    """A tracking number that landed on both a Naver and a 자연이랑 order is never included automatically."""
    channels: dict[str, set[str]] = {}
    for a in assignments:
        for tracking in a.tracking.split(", "):
            channels.setdefault(tracking, set()).add(channel_of[a.product_order_ids[0]])
    shared = {t for t, names in channels.items() if len(names) > 1}
    return tuple(
        replace(a, status=STATUS_CHECK, includable=False, reason="네이버·자연이랑 주문에 걸친 송장")
        if shared.intersection(a.tracking.split(", ")) else a for a in assignments)


def _split_channels(result: MatchResult, channel_of: dict[str, str]) -> tuple[MatchResult, tuple[Assignment, ...], int]:
    """(result with Naver orders only, the 자연이랑 assignments, 자연이랑 orders without a tracking number)."""
    def channel(ids: tuple[str, ...]) -> str:
        names = {channel_of[i] for i in ids}
        if len(names) > 1:
            raise ValueError("네이버와 자연이랑 주문이 한 묶음으로 합쳐졌습니다. 개발자에게 알려 주세요.")
        return names.pop()

    marked = _flag_cross_channel(result.assignments, channel_of)
    naver = tuple(a for a in marked if channel(a.product_order_ids) == NAVER_CHANNEL)
    withnature = tuple(a for a in marked if channel(a.product_order_ids) == CHANNEL)
    unshipped: tuple[UnshippedOrder, ...] = tuple(
        o for o in result.orders_without_invoice if channel(o.product_order_ids) == NAVER_CHANNEL)
    left = len(result.orders_without_invoice) - len(unshipped)
    return replace(result, assignments=naver, orders_without_invoice=unshipped), withnature, left


def _portal_rows(assignments: tuple[Assignment, ...], uploaded: dict[str, str]
                 ) -> tuple[list[tuple[str, str, str]], int, list[str]]:
    """(rows for the upload file, pairs skipped because already uploaded, shipments with an unknown courier)."""
    rows: list[tuple[str, str, str]] = []
    skipped, unknown = 0, []
    for a in assignments:
        code = courier_code(a.carrier)
        if code is None:
            unknown.append(f"{a.recipient} · {a.tracking} · {a.carrier}")
            continue
        for order_id in a.product_order_ids:
            if f"{order_id}|{a.tracking}" in uploaded:
                skipped += 1
            else:
                rows.append((order_id, a.tracking, code))
    return rows, skipped, unknown


def _match_both_channels(combined: pd.DataFrame, invoices, ledger: Path | None, moment: datetime
                         ) -> tuple[MatchResult, WithnatureShipments, list[InvoiceFile]]:
    channel_of = dict(zip(combined[ID_COLUMN], combined[CHANNEL_COLUMN], strict=True))
    result, withnature, unshipped = _split_channels(match_invoices(combined, invoices), channel_of)
    uploaded = read_withnature_ledger(ledger)
    matched = tuple(a for a in withnature if a.status != STATUS_CHECK)
    checks = tuple(a for a in withnature if a.status == STATUS_CHECK and a.includable)
    match_rows, skipped, unknown = _portal_rows(matched, uploaded)
    check_rows, skipped_checks, unknown_checks = _portal_rows(checks, uploaded)
    files = []
    if match_rows:
        files.append(InvoiceFile(WITHNATURE_FILE_NAME.format(when=moment), build_portal_upload(match_rows)))
        mark_withnature_uploaded([(order, tracking) for order, tracking, _ in match_rows], moment, ledger)
    if check_rows:
        files.append(InvoiceFile(WITHNATURE_CHECKS_FILE_NAME.format(when=moment), build_portal_upload(check_rows)))
    excluded = tuple(f"{a.recipient} · {a.reason}" for a in withnature if a.status == STATUS_CHECK and not a.includable)
    shipments = WithnatureShipments(len(match_rows), len(check_rows), skipped + skipped_checks,
                                    tuple(unknown + unknown_checks), excluded, unshipped)
    return result, shipments, files


# ---------------------------------------------------------------- run

def prepare_invoices(invoice_bytes: bytes, fetch: Callable[[], list[dict]] = fetch_paid_items,
                     ledger: Path = LEDGER_FILE, when: datetime | None = None,
                     withnature_orders: Path | None = None, withnature_ledger: Path | None = None) -> InvoiceRun:
    """Match the courier file against the unshipped orders and build the upload files.

    Saved 자연이랑 orders are matched in the same pass; their shipments go to a portal upload file and
    never into the Naver result, files or buttons."""
    invoices = read_invoice_file(invoice_bytes)
    shippable, _claims = split_shippable(fetch())
    table = normalize_orders(details_to_dataframe(shippable))
    done = set(read_ledger(ledger))
    fresh = table[~table[ID_COLUMN].isin(done)]
    saved = load_saved_orders(withnature_orders)
    moment = when or datetime.now(KST)
    if saved.empty:
        result, shipments, portal_files = match_invoices(fresh, invoices), None, []
    else:
        combined = _with_channels(fresh, saved)
        result, shipments, portal_files = _match_both_channels(combined, invoices, withnature_ledger, moment)
    files = []
    if result.matched():
        files.append(InvoiceFile(MATCHED_FILE_NAME.format(when=moment),
                                 build_upload_workbook(fresh, result.matched())))
    if result.includable_checks():
        files.append(InvoiceFile(WITH_CHECKS_FILE_NAME.format(when=moment),
                                 build_upload_workbook(fresh, result.includable())))
    return InvoiceRun(result, tuple(files + portal_files), len(table) - len(fresh), moment, shipments)


def register(entry: dict, include_checks: bool,
             dispatch: Callable[[list[DispatchItem]], DispatchResult] = dispatch_or_simulate,
             ledger: Path = LEDGER_FILE, now: Callable[[], datetime] = lambda: datetime.now(KST)) -> DispatchResult:
    """Send the pending run's tracking numbers to Naver; record the really registered ones."""
    done = set(read_ledger(ledger))
    groups = entry["includable" if include_checks else "matched"]
    # Naver product order ids are digits; a 자연이랑 주문번호 ('O26...') must never be sent to Naver
    items = {pid: DispatchItem(pid, tracking, courier)
             for ids, tracking, courier in map(_read_group, groups) for pid in ids
             if pid not in done and pid.isdigit()}
    tracking_of = {pid: item.tracking for pid, item in items.items()}
    result = dispatch(list(items.values()))
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


def _listed(items: tuple[str, ...]) -> list[str]:
    shown = [f"  - {item}" for item in items[:MAX_LISTED]]
    return shown + ([f"  … 외 {len(items) - MAX_LISTED}건"] if len(items) > MAX_LISTED else [])


def _withnature_lines(shipments: WithnatureShipments | None) -> list[str]:
    """The 자연이랑 part of the message; empty when no saved 자연이랑 order was involved."""
    if shipments is None or not (shipments.matched or shipments.checks or shipments.skipped
                                 or shipments.unknown_courier or shipments.excluded):
        return []
    lines = ["", f"🌿 {CHANNEL}"]
    if shipments.matched:
        lines.append(f"· {shipments.matched}건 → 파일을 사이트 '출고 및 운송장등록 > 일괄등록'에 올리고 "
                     "'첫 행은 컬럼으로 제외'를 체크하세요.")
    if shipments.checks:
        lines.append(f"· 확인 필요 {shipments.checks}건은 따로 파일(확인필요)에 담았습니다. 맞는지 보고 올려 주세요.")
    if shipments.skipped:
        lines.append(f"· 이미 올린 송장 {shipments.skipped}건 제외")
    if shipments.excluded:
        lines += [f"· ⚠️ 확인 필요(파일에서 제외) {len(shipments.excluded)}건"] + _listed(shipments.excluded)
    if shipments.unknown_courier:
        lines += [f"· ⚠️ 택배사를 몰라 뺀 송장 {len(shipments.unknown_courier)}건 (직접 등록)"] + _listed(shipments.unknown_courier)
    if shipments.unshipped:
        lines.append(f"· 송장 없는 {CHANNEL} 주문 {shipments.unshipped}건 (저장된 30일치)")
    return lines


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
    lines += _withnature_lines(run.withnature)
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


def dispatch_report(result: DispatchResult) -> str:
    """What the owner reads after tapping a 송장 등록 button."""
    lines = []
    if result.simulated:
        lines.append("🧪 시험 모드: 네이버에 실제로 등록하지 않았습니다 (INVOICE_DISPATCH_ENABLED 꺼짐).")
    done = "등록했을 송장" if result.simulated else "등록 완료 송장"
    lines.append(f"✅ {done} {len(result.success_ids)}건")
    if result.failures:
        lines.append(f"⚠️ 실패 {len(result.failures)}건")
        lines += [f"· {f.id} · {f.message or f.code}" for f in result.failures[:MAX_FAILURE_LINES]]
        if len(result.failures) > MAX_FAILURE_LINES:
            lines.append(f"… 그 외 {len(result.failures) - MAX_FAILURE_LINES}건")
    return "\n".join(lines)
