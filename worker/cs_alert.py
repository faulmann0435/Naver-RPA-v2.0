"""Telegram text for one new customer question: question, linked order, AI draft (plain text).

Only non-personal order facts are shown — OrderSummary already leaves out names, phones, addresses,
order ids and tracking numbers, and nothing else is read from the raw data here.
"""
from __future__ import annotations

from worker.cs_drafter import ROUTE_AI, DraftResult, InquiryView, stamp_with_weekday
from worker.order_link import KIND_GUESS, KIND_MANY, LinkResult, OrderSummary

TELEGRAM_LIMIT = 4096
SAFE_LIMIT = 4000  # leave room below Telegram's hard limit
PRODUCT_MAX, QUESTION_MAX, OPTION_MAX, REASON_MAX = 30, 300, 20, 100
ICONS = {"상품문의": "🟡", "고객문의": "🔵"}
CLAIM_ICON, CLAIM_KIND = "🔴", "불만·클레임"
NO_AI_NOTE = "(AI 초안 꺼짐 — ANTHROPIC_API_KEY 없음)"
NO_DRAFT = "(초안 없음)"
CHECK_HINT = "→ 판매자센터에서 작성자 아이디를 눌러 실제 주문을 확인해 주세요"


def trim(text: str, limit: int) -> str:
    """Text cut to `limit` characters, ending with '…' when something was removed."""
    flat = (text or "").strip()
    return flat if len(flat) <= limit else flat[:max(limit - 1, 0)].rstrip() + "…"


def _order_line(order: OrderSummary) -> str:
    option = trim(order.option, OPTION_MAX)
    parts = [order.paid_at.strftime("%Y-%m-%d %H:%M") if order.paid_at else "",
             order.status_label,
             f"{option} ×{order.quantity}" if option else f"×{order.quantity}"]
    return "- " + " · ".join(p for p in parts if p)


def _order_block(link: LinkResult) -> str:
    lines = [f"주문: {link.kind}"] + [_order_line(o) for o in link.orders]
    if link.kind in (KIND_GUESS, KIND_MANY):
        lines.append(CHECK_HINT)
    return "\n".join(lines)


def _icon(view: InquiryView, draft: DraftResult | None) -> str:
    if draft is not None and draft.kind == CLAIM_KIND:
        return CLAIM_ICON
    return ICONS.get(view.channel, "🟡")


def _class_line(draft: DraftResult) -> str:
    verdict = "AI 초안" if draft.route == ROUTE_AI else f"사장님 직접 처리 권장 — {trim(draft.reason, REASON_MAX)}"
    return f"분류: {draft.kind} · {verdict}"


def format_alert(view: InquiryView, link_result: LinkResult, draft: DraftResult | None,
                 ai_note: str = NO_AI_NOTE) -> str:
    """The alert text; `ai_note` replaces the 분류/초안 part when there is no draft (off or failed)."""
    head = "\n".join([
        f"{_icon(view, draft)} {view.channel} · {trim(view.product_name, PRODUCT_MAX)}",
        stamp_with_weekday(view.created),
        f'고객: "{trim(view.question, QUESTION_MAX)}"',
        _order_block(link_result),
    ])
    if draft is None:
        return trim(f"{head}\n{ai_note}", TELEGRAM_LIMIT)
    top = f"{head}\n{_class_line(draft)}\n✍️ AI 초안:\n"
    body = trim(draft.draft, SAFE_LIMIT - len(top)) if draft.draft else NO_DRAFT
    return top + body
