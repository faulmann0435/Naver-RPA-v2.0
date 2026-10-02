"""Normalized option keys and product numbers used to look up the item dictionary."""
from __future__ import annotations

import math
import re
import unicodedata
from collections.abc import Sequence
from numbers import Integral

OPTION_SEPARATOR = " / "
GROUP_VALUE_SEPARATOR = ": "
_KEEP_CODEPOINTS_REMOVED = frozenset({0xFE0F, 0xFE0E, 0x200D})  # variation selectors, ZWJ
_MODIFIER_RANGE = range(0x1F3FB, 0x1F400)  # emoji skin-tone modifiers
_WHITESPACE = re.compile(r"\s+")
_TRAILING_ZERO = re.compile(r"^(\d+)\.0$")


def _is_missing(value: object) -> bool:
    return value is None or (isinstance(value, float) and math.isnan(value))


def _is_decoration(ch: str) -> bool:
    category = unicodedata.category(ch)
    code = ord(ch)
    if category == "So":
        return True
    if category in ("Mn", "Cf") and code in _KEEP_CODEPOINTS_REMOVED:
        return True
    return category == "Sk" and code in _MODIFIER_RANGE


def normalize_text(text: str) -> str:
    """NFKC, drop emoji/symbol decorations, collapse whitespace."""
    folded = unicodedata.normalize("NFKC", text)
    stripped = "".join(ch for ch in folded if not _is_decoration(ch))
    return _WHITESPACE.sub(" ", stripped).strip()


def make_option_key(option_text: object, ignored_groups: Sequence[str] = ()) -> str:
    """Normalize an option string; drop segments whose group is in `ignored_groups`."""
    if _is_missing(option_text):
        return ""
    ignored = {g for g in (normalize_text(str(x)) for x in ignored_groups) if g}
    kept: list[str] = []
    for segment in str(option_text).split(OPTION_SEPARATOR):
        group, found, _ = segment.partition(GROUP_VALUE_SEPARATOR)
        if found and normalize_text(group) in ignored:
            continue
        normalized = normalize_text(segment)
        if normalized:
            kept.append(normalized)
    return OPTION_SEPARATOR.join(kept)


def normalize_product_no(value: object) -> str:
    """Product number as a plain digit string (5568579375, 5568579375.0, ' 5568579375 ')."""
    if _is_missing(value):
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, Integral):
        return str(int(value))
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    text = str(value).strip()
    match = _TRAILING_ZERO.match(text)
    return match.group(1) if match else text
