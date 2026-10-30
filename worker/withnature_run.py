"""One 자연이랑 run: the portal's 출고작업관리 Excel -> purchase-order files, plus a saved copy of its orders.

run_withnature() turns the export into the Naver-shaped table, runs the same rules, layouts and item
dictionary as [📦 발주 처리], and renames each file 자연이랑_<name>. Nothing is sent anywhere.
The converted rows are kept under DATA_DIR/발주기록/자연이랑_주문.json (30 days), so a tracking file
that arrives later can be matched to them (worker.invoice_run).

    {"O26102986789": {"saved": "2026-10-30", "row": {"수취인명": ..., ...}}, ...}
"""
from __future__ import annotations

import json
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

from core.pipeline import ProcessResult, process_orders
from core.withnature import (
    CHANNEL,
    ORDER_COLUMNS,
    read_withnature_export,
    to_order_frame,
)
from worker.naver_adapter import KST
from worker.order_run import Engine, _names, load_engine
from worker.paths import DATA_DIR

ORDERS_FILE = DATA_DIR / "발주기록" / "자연이랑_주문.json"
KEEP_DAYS = 30
FILE_PREFIX = f"{CHANNEL}_"
ID_COLUMN = "상품주문번호"
NAME_COLUMN = "수취인명"


@dataclass(frozen=True)
class WithnatureRun:
    files: list[dict]
    order_count: int
    vendor_counts: tuple[tuple[str, int], ...]  # (file name, orders on it)
    unmatched: int
    unclassified: int
    needs_review: int
    source: str
    repeated_names: tuple[str, ...] = ()  # recipients whose 주문번호 was already saved by an earlier run
    when: datetime = field(default_factory=lambda: datetime.now(KST))


# ---------------------------------------------------------------- saved orders

def _json_value(value):
    if hasattr(value, "item"):
        return value.item()
    return value


def _read_store(path: Path) -> dict[str, dict]:
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


def _fresh(store: dict[str, dict], today: date) -> dict[str, dict]:
    """The entries saved within the last KEEP_DAYS days."""
    oldest = today - timedelta(days=KEEP_DAYS)
    return {k: v for k, v in store.items() if date.fromisoformat(v["saved"]) >= oldest}


def save_orders(frame: pd.DataFrame, today: date, path: Path | None = None) -> None:
    """Merge the converted rows into the store (same 주문번호: the newer row wins), dropping old entries."""
    path = path or ORDERS_FILE
    merged = _fresh(_read_store(path), today)
    for row in frame.to_dict("records"):
        merged[str(row[ID_COLUMN])] = {"saved": today.isoformat(), "row": {k: _json_value(v) for k, v in row.items()}}
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(merged, ensure_ascii=False, indent=1), encoding="utf-8")


def load_saved_orders(path: Path | None = None, today: date | None = None) -> pd.DataFrame:
    """The saved 자연이랑 orders (not older than KEEP_DAYS) as the Naver-shaped table; empty when none."""
    store = _fresh(_read_store(path or ORDERS_FILE), today or datetime.now(KST).date())
    rows = [entry["row"] for entry in store.values()]
    return pd.DataFrame(rows, columns=list(ORDER_COLUMNS)) if rows else pd.DataFrame(columns=list(ORDER_COLUMNS))


# ---------------------------------------------------------------- run

def _renamed(files: list[dict]) -> list[dict]:
    return [{**f, "filename": FILE_PREFIX + f["filename"]} for f in files]


def _vendor_counts(result: ProcessResult) -> tuple[tuple[str, int], ...]:
    merged = result.merged
    if merged is None or merged.empty:
        return ()
    counts = merged["_VendorID"].astype(str).str.strip().value_counts()
    names = {f["vendor"]: f["filename"] for f in _renamed(result.files)}
    return tuple((names[v], int(n)) for v, n in counts.items() if v in names)


def run_withnature(xlsx_bytes: bytes, engine_loader: Callable[[], Engine] = load_engine,
                   orders_file: Path | None = None, now: Callable[[], datetime] = lambda: datetime.now(KST)) -> WithnatureRun:
    """Convert, build the purchase orders, save the orders. Changes nothing on the data store."""
    frame = to_order_frame(read_withnature_export(xlsx_bytes))
    when = now()
    path = orders_file or ORDERS_FILE
    earlier = _fresh(_read_store(path), when.date())
    repeated = tuple(dict.fromkeys(str(r[NAME_COLUMN]) or "이름 없음" for r in frame.to_dict("records")
                                   if str(r[ID_COLUMN]) in earlier))
    engine = engine_loader()
    result = process_orders(frame, engine.config, engine.dictionary, engine.settings)
    save_orders(frame, when.date(), path)
    return WithnatureRun(
        files=_renamed(result.files), order_count=len(frame), vendor_counts=_vendor_counts(result),
        unmatched=len(result.unmatched), unclassified=int(result.stats.get("rows_unclassified", 0)),
        needs_review=int(result.stats.get("rows_needs_review", 0)), source=engine.source,
        repeated_names=repeated, when=when)


# ---------------------------------------------------------------- message

def summary(run: WithnatureRun) -> str:
    """The Korean message that goes with the files."""
    lines = [f"🌿 자연이랑 발주서 {len(run.files)}개 ({run.when:%m/%d %H:%M})",
             f"· 주문 {run.order_count}건"]
    lines += [f"  - {name}: {count}건" for name, count in run.vendor_counts]
    if run.unclassified:
        lines.append(f"· ⚠️ 업체를 못 정한 주문 {run.unclassified}건 (발주서에 없음) → 앱의 품목 사전에서 등록해 주세요.")
    if run.unmatched:
        lines.append(f"· ⚠️ 미등록 품목 {run.unmatched}건 → 앱에서 등록한 뒤 파일을 다시 보내 주세요 "
                     f"(상품번호가 '{CHANNEL}-'로 시작합니다).")
    else:
        lines.append("· ✅ 미등록 품목 없음")
    if run.needs_review:
        lines.append(f"· ℹ️ 사전 검토 표시가 남은 품목 {run.needs_review}줄 (발주서에는 정상 반영, 사전 정리용 참고)")
    if run.repeated_names:
        lines.append(f"· ⚠️ 전에 받은 파일에도 있던 주문 {len(run.repeated_names)}건: {_names(run.repeated_names)}\n"
                     "  이미 앞선 발주서에 넣어 현장에 보낸 주문이면 두 번 보내지 않게 확인하세요.")
    return "\n".join(lines)
