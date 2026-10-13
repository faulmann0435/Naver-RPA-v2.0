"""송장 입력 page logic: Naver order export + Hanjin file -> matching tables and the upload xlsx.

Pure functions (no streamlit); the page in ui/invoice_page.py only shows what they return.
"""
from __future__ import annotations

import hashlib
from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd

from core.invoice_match import (
    COURIER_CODES,
    KST,
    Assignment,
    MatchResult,
    build_upload_workbook,
    match_invoices,
    read_invoice_file,
    read_order_export,
)
from core.loader import _get_excel_bytes

UPLOAD_FILE_NAME = "송장업로드_{when:%Y%m%d_%H%M}.xlsx"
REGISTER_SENT = "요청을 보냈습니다. 사장님 PC가 30초 안에 처리하고 결과를 여기와 텔레그램에 알려 드립니다."
REGISTER_NEEDS_STORE = "네이버에 바로 등록하려면 데이터 저장소 설정이 필요합니다. 지금은 위의 📥 파일로 올려 주세요."
STALE_AFTER = timedelta(minutes=3)
STALE_WARNING = "사장님 PC가 꺼져 있거나 봇이 멈췄을 수 있습니다. 위의 📥 파일을 내려받아 직접 올려 주세요."
NOTE_INSTRUCTION = ("내려받은 파일을 스마트스토어센터 판매관리 > 발주(주문)확인/발송관리 > "
                    "엑셀 일괄발송에서 올리세요.")


@dataclass(frozen=True)
class InvoiceOutcome:
    orders: pd.DataFrame
    result: MatchResult


def match_files(order_bytes: bytes, invoice_bytes: bytes | list[bytes],
                password: str | None = None) -> InvoiceOutcome:
    """Decrypt and read the order export and every tracking file, then match. Raises ValueError (Korean)."""
    plain = _get_excel_bytes(BytesIO(order_bytes), password=password)
    orders = read_order_export(plain)
    files = [invoice_bytes] if isinstance(invoice_bytes, bytes) else invoice_bytes
    invoices = [invoice for data in files for invoice in read_invoice_file(data)]
    return InvoiceOutcome(orders, match_invoices(orders, invoices))


def metrics(result: MatchResult) -> dict[str, int]:
    return {
        "일치": len(result.matched()),
        "확인 필요": len(result.needs_check()),
        "주문 못 찾은 송장": len(result.invoices_without_order),
        "송장 없는 주문": len(result.orders_without_invoice),
    }


def _assignment_rows(assignments: tuple[Assignment, ...]) -> list[dict]:
    return [{"수취인": a.recipient, "상품": a.items, "송장번호": a.tracking,
             "상품주문번호": ", ".join(a.product_order_ids)} for a in assignments]


def matched_table(result: MatchResult) -> pd.DataFrame:
    return pd.DataFrame(_assignment_rows(result.matched()), columns=["수취인", "상품", "송장번호", "상품주문번호"])


def check_table(result: MatchResult) -> pd.DataFrame:
    rows = [{**row, "사유": a.reason, "포함 가능": "예" if a.includable else "아니오"}
            for a, row in zip(result.needs_check(), _assignment_rows(result.needs_check()), strict=True)]
    return pd.DataFrame(rows, columns=["수취인", "사유", "포함 가능", "송장번호", "상품", "상품주문번호"])


def loose_invoice_table(result: MatchResult) -> pd.DataFrame:
    rows = [{"순번": i.seq, "받는분": i.name, "송장번호": i.tracking, "품목": i.item_name}
            for i in result.invoices_without_order]
    return pd.DataFrame(rows, columns=["순번", "받는분", "송장번호", "품목"])


def no_tracking_table(result: MatchResult) -> pd.DataFrame:
    rows = [{"줄": i.seq, "받는분": i.name or "(비어 있음)", "품목": i.item_name}
            for i in result.invoices_without_tracking]
    return pd.DataFrame(rows, columns=["줄", "받는분", "품목"])


def unshipped_table(result: MatchResult) -> pd.DataFrame:
    rows = [{"수취인": o.recipient, "상품": o.items, "상품주문번호": ", ".join(o.product_order_ids)}
            for o in result.orders_without_invoice]
    return pd.DataFrame(rows, columns=["수취인", "상품", "상품주문번호"])


def build_download(outcome: InvoiceOutcome, include_checks: bool, when: datetime | None = None) -> tuple[str, bytes]:
    """(file name, xlsx bytes): matched rows, plus the includable checks when asked."""
    assignments = outcome.result.includable() if include_checks else outcome.result.matched()
    moment = when or datetime.now(KST)
    return UPLOAD_FILE_NAME.format(when=moment), build_upload_workbook(outcome.orders, assignments)


# ---------------------------------------------------------------- 네이버에 바로 등록 (queue to the owner's PC)

def queue_items(assignments: tuple[Assignment, ...]) -> list[dict]:
    """Shipments in the dispatch queue format (store.dispatch_queue)."""
    return [{"ids": list(a.product_order_ids), "tracking": a.tracking, "courier": COURIER_CODES[a.carrier],
             "recipient": a.recipient} for a in assignments]


def selection_key(items: list[dict]) -> str:
    """Fingerprint of a selection, so the same one is not submitted twice from this session."""
    pairs = sorted(f"{i}:{item['tracking']}" for item in items for i in item["ids"])
    return hashlib.sha1(",".join(pairs).encode()).hexdigest()[:16]


def status_messages(request: dict, now: datetime) -> list[tuple[str, str]]:
    """(level, Korean text) lines for a queued request; level is info, warning, success or error."""
    status, result = request.get("status"), request.get("result") or {}
    if status == "pending":
        lines = [("info", "⏳ 사장님 PC 처리 대기 중…")]
        created = _parse_time(request.get("created_at"))
        if created is not None and now - created > STALE_AFTER:
            lines.append(("warning", STALE_WARNING))
        return lines
    if status == "failed":
        return [("error", f"⚠️ 등록하지 못했습니다: {result.get('error') or '알 수 없는 오류'}")]
    if status != "done":
        return [("warning", "요청 상태를 알 수 없습니다.")]
    done, failures = len(result.get("success_ids", [])), result.get("failures", [])
    lines = []
    if result.get("simulated"):
        lines.append(("warning", "🧪 시험 모드: 실제로 등록되지 않았습니다"))
    lines.append(("success", f"✅ {'등록했을 송장' if result.get('simulated') else '등록 완료'} {done}건"))
    if failures:
        lines.append(("error", f"⚠️ 실패 {len(failures)}건"))
        lines += [("error", f"· {f.get('id')} · {f.get('message') or f.get('code')}") for f in failures]
    return lines


def _parse_time(value: object) -> datetime | None:
    try:
        moment = datetime.fromisoformat(str(value))
    except ValueError:
        return None
    return moment if moment.tzinfo else moment.replace(tzinfo=KST)
