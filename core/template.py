"""Display templates: literal text with {수량} and {수량*N} placeholders."""
from __future__ import annotations

import re

_PLACEHOLDER = re.compile(r"\{([^{}]*)\}")
_QTY_BODY = re.compile(r"수량(?:\s*\*\s*(\d+(?:\.\d+)?))?")


class TemplateError(ValueError):
    """The template text is not valid."""


def validate_template(t: str) -> list[str]:
    """Return Korean error messages; an empty list means the template is valid."""
    errors: list[str] = []
    for match in _PLACEHOLDER.finditer(t):
        if not _QTY_BODY.fullmatch(match.group(1).strip()):
            errors.append(f"알 수 없는 자리표시자입니다: {match.group(0)} ({{수량}} 또는 {{수량*숫자}}만 쓸 수 있습니다)")
    rest = _PLACEHOLDER.sub("", t)
    if "{" in rest or "}" in rest:
        errors.append("중괄호 { } 의 짝이 맞지 않습니다.")
    return errors


def has_qty_placeholder(t: str) -> bool:
    return any(_QTY_BODY.fullmatch(m.group(1).strip()) for m in _PLACEHOLDER.finditer(t))


def format_number(x: float) -> str:
    """Integer-valued -> no decimals; otherwise up to 3 decimals, trailing zeros stripped."""
    rounded = round(float(x), 3)
    if rounded == int(rounded):
        return str(int(rounded))
    return f"{rounded:.3f}".rstrip("0").rstrip(".")


def render(t: str, qty: int) -> str:
    errors = validate_template(t)
    if errors:
        raise TemplateError("; ".join(errors))

    def _replace(match: re.Match[str]) -> str:
        body = _QTY_BODY.fullmatch(match.group(1).strip())
        factor = body.group(1) if body else None
        return format_number(qty * float(factor) if factor else qty)

    return _PLACEHOLDER.sub(_replace, t)
