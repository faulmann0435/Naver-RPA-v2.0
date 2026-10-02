"""Plain sentence <-> display template: tick which numbers scale with the quantity (no streamlit)."""
from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass

from core.template import validate_template

_NUMBER = re.compile(r"\d+(?:\.\d+)?")
_PLACEHOLDER = re.compile(r"\{([^{}]*)\}")
_QTY_BODY = re.compile(r"수량(?:\s*\*\s*(\d+(?:\.\d+)?))?")


@dataclass(frozen=True)
class Piece:
    text: str
    is_number: bool


def split_pieces(sentence: str) -> list[Piece]:
    """Alternating text / number pieces; empty text pieces are left out."""
    pieces: list[Piece] = []
    pos = 0
    for match in _NUMBER.finditer(sentence):
        if match.start() > pos:
            pieces.append(Piece(sentence[pos:match.start()], False))
        pieces.append(Piece(match.group(0), True))
        pos = match.end()
    if pos < len(sentence):
        pieces.append(Piece(sentence[pos:], False))
    return pieces


def numbers_of(sentence: str) -> list[str]:
    return [p.text for p in split_pieces(sentence) if p.is_number]


def _scaled(number: str) -> str:
    return "{수량}" if float(number) == 1 else "{수량*" + number + "}"


def build_template(sentence: str, scale: Sequence[bool]) -> str:
    """scale[i] applies to the i-th number of the sentence; missing entries mean False."""
    out: list[str] = []
    index = 0
    for piece in split_pieces(sentence):
        if piece.is_number:
            flag = index < len(scale) and bool(scale[index])
            out.append(_scaled(piece.text) if flag else piece.text)
            index += 1
        else:
            out.append(piece.text)
    return "".join(out)


def _canonical(template: str) -> str:
    def _replace(match: re.Match[str]) -> str:
        body = _QTY_BODY.fullmatch(match.group(1).strip())
        factor = body.group(1) if body else None
        return "{수량*" + factor + "}" if factor else "{수량}"

    return _PLACEHOLDER.sub(_replace, template)


def sentence_and_scale(template: str) -> tuple[str, list[bool]]:
    """Inverse of build_template. A template that cannot be shown as a sentence comes back unchanged."""
    if validate_template(template):
        return template, []
    sentence: list[str] = []
    scale: list[bool] = []
    pos = 0
    for match in _PLACEHOLDER.finditer(template):
        literal = template[pos:match.start()]
        sentence.append(literal)
        scale.extend([False] * len(_NUMBER.findall(literal)))
        body = _QTY_BODY.fullmatch(match.group(1).strip())
        factor = body.group(1) if body else None
        sentence.append(factor or "1")
        scale.append(True)
        pos = match.end()
    tail = template[pos:]
    sentence.append(tail)
    scale.extend([False] * len(_NUMBER.findall(tail)))
    text = "".join(sentence)
    # Digits next to a placeholder would merge into one number (e.g. "{수량}5"): keep the raw template.
    if len(numbers_of(text)) != len(scale) or build_template(text, scale) != _canonical(template):
        return template, []
    return text, scale


def guess_scale(old_sentence: str, old_scale: Sequence[bool], new_sentence: str) -> list[bool]:
    """Flags for an edited sentence: by position when the count is unchanged, else by number value."""
    new_numbers = numbers_of(new_sentence)
    old_numbers = numbers_of(old_sentence)
    flags = [i < len(old_scale) and bool(old_scale[i]) for i in range(len(old_numbers))]
    if len(new_numbers) == len(old_numbers):
        return flags
    scaled_values = {float(n) for n, flag in zip(old_numbers, flags, strict=True) if flag}
    return [float(n) in scaled_values for n in new_numbers]
