"""송장 등록 기록 page logic (pure): labels and tables over the entries of dispatch/queue.json."""
from __future__ import annotations

from collections.abc import Mapping, Sequence
from datetime import datetime

import pandas as pd

from core.invoice_match import COURIER_CODES, KST
from store import dispatch_queue as queue

FILTER_ALL, FILTER_TELEGRAM, FILTER_APP = "전체", "텔레그램", "사이트"
FILTER_CHOICES = (FILTER_ALL, FILTER_TELEGRAM, FILTER_APP)
_FILTER_SOURCE = {FILTER_TELEGRAM: queue.SOURCE_TELEGRAM, FILTER_APP: queue.SOURCE_APP}
COURIER_NAMES = {code: name for name, code in COURIER_CODES.items()}
RESULT_OK, RESULT_SIMULATED, RESULT_WAITING, RESULT_UNKNOWN = "등록됨", "시험", "대기", "–"
SUMMARY_COLUMNS = ["시간", "경로", "요청자", "상태", "송장 수", "성공", "실패"]
DETAIL_COLUMNS = ["받는 사람", "송장번호", "택배사", "결과", "상품주문번호"]


def _result(request: Mapping) -> Mapping:
    return request.get("result") or {}


def source_label(request: Mapping) -> str:
    return "📱 텔레그램" if queue.source_of(request) == queue.SOURCE_TELEGRAM else "🖥️ 사이트"


def status_label(request: Mapping) -> str:
    status, result = request.get("status"), _result(request)
    if status == queue.STATUS_PENDING:
        return "⏳ 처리 대기"
    if status == queue.STATUS_FAILED:
        return "❌ 실패"
    if result.get("simulated"):
        return "🧪 시험 모드 (실제 등록 안 됨)"
    if result.get("failures"):
        return "⚠️ 일부 실패"
    return "✅ 완료"


def time_label(request: Mapping) -> str:
    """created_at as MM/DD HH:MM in KST (the raw text when it cannot be read)."""
    raw = str(request.get("created_at", ""))
    try:
        moment = datetime.fromisoformat(raw)
    except ValueError:
        return raw
    moment = moment.replace(tzinfo=KST) if moment.tzinfo is None else moment.astimezone(KST)
    return f"{moment:%m/%d %H:%M}"


def filter_by_source(requests: Sequence[Mapping], choice: str) -> list[Mapping]:
    wanted = _FILTER_SOURCE.get(choice)
    return [r for r in requests if wanted is None or queue.source_of(r) == wanted]


def summary_table(requests: Sequence[Mapping]) -> pd.DataFrame:
    rows = [{"시간": time_label(r), "경로": source_label(r), "요청자": str(r.get("by", "")),
             "상태": status_label(r), "송장 수": len(r.get("items") or []),
             "성공": len(_result(r).get("success_ids") or []), "실패": len(_result(r).get("failures") or [])}
            for r in requests]
    return pd.DataFrame(rows, columns=SUMMARY_COLUMNS)


def _shipment_result(request: Mapping, item: Mapping) -> str:
    result = _result(request)
    ids = set(item.get("ids") or [])
    messages = [str(f.get("message") or f.get("code") or "") for f in result.get("failures") or []
                if f.get("id") in ids]
    if request.get("status") == queue.STATUS_PENDING:
        return RESULT_WAITING
    if messages:
        return f"실패: {messages[0]}"
    if ids and ids <= set(result.get("success_ids") or []):
        return RESULT_SIMULATED if result.get("simulated") else RESULT_OK
    if request.get("status") == queue.STATUS_FAILED and result.get("error"):
        return f"실패: {result['error']}"
    return RESULT_UNKNOWN


def detail_table(request: Mapping) -> pd.DataFrame:
    rows = [{"받는 사람": str(i.get("recipient", "")), "송장번호": str(i.get("tracking", "")),
             "택배사": COURIER_NAMES.get(i.get("courier"), str(i.get("courier", ""))),
             "결과": _shipment_result(request, i), "상품주문번호": ", ".join(i.get("ids") or [])}
            for i in request.get("items") or []]
    return pd.DataFrame(rows, columns=DETAIL_COLUMNS)


def expander_title(request: Mapping) -> str:
    return (f"{time_label(request)} · {source_label(request)} · {status_label(request)} · "
            f"{len(request.get('items') or [])}건")
