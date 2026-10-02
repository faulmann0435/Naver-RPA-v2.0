"""Draft item-dictionary seeding: probe the option rules, derive templates from hand text (pure logic)."""
from __future__ import annotations

import re
from collections import Counter
from collections.abc import Sequence
from dataclasses import dataclass, replace

import pandas as pd

from core.actions import _safe_int
from core.dictionary import DEFAULT_SEPARATOR, DictionaryEntry
from core.engine import apply_option_rules
from core.merger import _cleanup_empty_parens, _dictionary_parts, _format_weight
from core.option_key import GROUP_VALUE_SEPARATOR, OPTION_SEPARATOR, normalize_text
from core.template import TemplateError, format_number, render, validate_template

SAMPLE_QTYS: tuple[int, ...] = (1, 2, 3)
TOLERANCE = 1e-9
_NUMBER = re.compile(r"(\d+(?:\.\d+)?)")


@dataclass(frozen=True)
class Sample:
    """Engine result for one quantity: (text, weight in kg, quantity already stamped in text)."""

    text: str
    weight: float
    formatted: bool


@dataclass(frozen=True)
class Probe:
    """What the option rules do to one (product, option): a template, or a weight group."""

    template: str
    sum_group: str
    unit_weight_kg: float | None
    failed: bool
    samples: tuple[Sample, ...]
    q1_differs: bool = False  # engine output for qty 1 differs in shape (stray 'x1' stamp); template uses qty 2-3


@dataclass(frozen=True)
class HandEvidence:
    """A hand-written text for a single-row bundle: `hand` was written for `qty` units."""

    hand: str
    qty: int
    option_text: str


@dataclass(frozen=True)
class EvidenceResult:
    entry: DictionaryEntry
    used: bool
    conflict: bool
    rejected: int


def _close(a: float, b: float) -> bool:
    return abs(a - b) <= TOLERANCE


def strip_ignored_groups(option_text: str, ignored_groups: Sequence[str]) -> str:
    """Drop the option segments whose group is ignored (as in the key), keeping the rest untouched."""
    ignored = {g for g in (normalize_text(x) for x in ignored_groups) if g}
    kept: list[str] = []
    for segment in str(option_text).split(OPTION_SEPARATOR):
        group, found, _ = segment.partition(GROUP_VALUE_SEPARATOR)
        if found and normalize_text(group) in ignored:
            continue
        kept.append(segment)
    return OPTION_SEPARATOR.join(kept)


def run_samples(
    product_name: str,
    option_text: str,
    vendor: str,
    option_rules: pd.DataFrame,
    qtys: Sequence[int] = SAMPLE_QTYS,
) -> tuple[Sample, ...]:
    """Run the option engine on a one-row order for each quantity."""
    samples: list[Sample] = []
    for qty in qtys:
        row = pd.Series({"상품명": product_name, "옵션정보": option_text, "수량": qty, "_VendorID": vendor})
        text, weight, formatted = apply_option_rules(row, option_rules)
        samples.append(Sample(str(text), float(weight), bool(formatted)))
    return tuple(samples)


def _template_from_samples(samples: Sequence[Sample], qtys: Sequence[int]) -> str | None:
    parts = [_NUMBER.split(s.text) for s in samples]
    if len({len(p) for p in parts}) != 1:
        return None
    skeletons = [p[0::2] for p in parts]
    if any(s != skeletons[0] for s in skeletons[1:]):
        return None
    out: list[str] = []
    for i, piece in enumerate(parts[0]):
        if i % 2 == 0:
            out.append(piece)
            continue
        numbers = [float(p[i]) for p in parts]
        first = numbers[0] / qtys[0]
        if all(_close(n, numbers[0]) for n in numbers):
            out.append(piece)
        elif all(_close(n, q * first) for n, q in zip(numbers, qtys)):
            out.append("{수량}" if _close(first, 1) else "{수량*" + format_number(first) + "}")
        else:
            return None
    template = "".join(out)
    return None if validate_template(template) else template


def probe_samples(samples: tuple[Sample, ...]) -> Probe:
    """Turn engine samples (qty 1, 2, 3) into a template or a weight group; `failed` when not linear."""
    first = samples[0]
    if first.weight > 0:
        linear = bool(first.text) and all(_close(s.weight, q * first.weight) for s, q in zip(samples, SAMPLE_QTYS))
        if linear:
            return Probe(first.text, first.text, first.weight, False, samples)
        return Probe(first.text, "", None, True, samples)
    template = _template_from_samples(samples, SAMPLE_QTYS)
    if template is not None:
        return Probe(template, "", None, False, samples)
    # Rules like FORMAT_QTY stamp ' x1' only when qty == 1 (CALC_UNIT changed nothing); try qty 2-3 alone.
    template = _template_from_samples(samples[1:], SAMPLE_QTYS[1:])
    if template is not None:
        return Probe(template, "", None, False, samples, q1_differs=True)
    return Probe(first.text, "", None, True, samples)


def probe_row(product_name: str, option_text: str, vendor: str, option_rules: pd.DataFrame) -> Probe:
    return probe_samples(run_samples(product_name, option_text, vendor, option_rules))


def probe_entry(
    probe: Probe,
    product_no: str,
    option_key: str,
    vendor_id: str,
    product_name_ref: str = "",
    option_raw_ref: str = "",
) -> DictionaryEntry:
    return DictionaryEntry(
        product_no=product_no,
        option_key=option_key,
        vendor_id=vendor_id,
        display_template=probe.template,
        sum_group=probe.sum_group,
        unit_weight_kg=probe.unit_weight_kg,
        needs_review=True,
        product_name_ref=product_name_ref,
        option_raw_ref=option_raw_ref,
    )


def entry_display(entry: DictionaryEntry, qty: int) -> str | None:
    """What the dictionary merge shows for a single-row bundle of `qty` units (None: invalid template)."""
    frame = pd.DataFrame({"_dict_entry": pd.Series([entry], dtype=object), "수량": [qty]})
    try:
        parts, suffixes = _dictionary_parts(frame)
    except TemplateError:
        return None
    text = DEFAULT_SEPARATOR.join(parts)
    if suffixes:
        text = (text + " " if text else "") + " ".join(suffixes)
    return str(_cleanup_empty_parens(text))


def expected_display(samples: Sequence[Sample], qty: int) -> str:
    """What the classic (rule engine + merge) path shows for a single-row bundle of `qty` units."""
    sample = samples[SAMPLE_QTYS.index(qty)]
    if sample.weight > 0:
        text = f"{sample.text} {_format_weight(sample.weight)}"
    elif not sample.text:
        text = ""
    elif sample.formatted or qty <= 1:
        text = sample.text
    else:
        text = f"{sample.text} (x{qty})"
    return str(_cleanup_empty_parens(text))


def verify_probe(entry: DictionaryEntry, probe: Probe) -> bool:
    """True when the entry reproduces the classic output for qty 1, 2 and 3 (qty 1 skipped if `q1_differs`)."""
    qtys = SAMPLE_QTYS[1:] if probe.q1_differs else SAMPLE_QTYS
    return all(entry_display(entry, q) == expected_display(probe.samples, q) for q in qtys)


def _number_piece(piece: str, qty: int, option_numbers: set[str]) -> str:
    value = float(piece)
    if qty <= 1:
        return piece
    if _close(value, qty):
        return "{수량}"
    if value.is_integer() and int(value) % qty == 0:
        factor = int(value) // qty
        if format_number(factor) in option_numbers:
            return "{수량*" + str(factor) + "}"
    return piece


def derive_template(hand: str, qty: int, option_text: str) -> str | None:
    """Template that renders back to `hand` for `qty` units; None when that is impossible."""
    option_numbers = {format_number(float(n)) for n in _NUMBER.findall(option_text)}
    parts = _NUMBER.split(hand)
    template = "".join(
        piece if i % 2 == 0 else _number_piece(piece, qty, option_numbers) for i, piece in enumerate(parts)
    )
    if validate_template(template):
        return None
    return template if render(template, qty) == hand else None


def _candidate(base: DictionaryEntry, evidence: HandEvidence) -> DictionaryEntry | None:
    if entry_display(base, evidence.qty) == evidence.hand:
        return base
    template = derive_template(evidence.hand, evidence.qty, evidence.option_text)
    if template is None:
        return None
    candidate = replace(base, display_template=template, display_template_qty1="", sum_group="", unit_weight_kg=None)
    return candidate if entry_display(candidate, evidence.qty) == evidence.hand else None


def apply_evidence(base: DictionaryEntry, evidence: Sequence[HandEvidence]) -> EvidenceResult:
    """Pick the entry the hand texts agree on (most frequent); the probe entry wins when it reproduces them."""
    candidates = [c for c in (_candidate(base, e) for e in evidence) if c is not None]
    rejected = len(evidence) - len(candidates)
    if not candidates:
        return EvidenceResult(base, False, False, rejected)
    counts = Counter(candidates)
    winner = counts.most_common(1)[0][0]
    return EvidenceResult(winner, True, len(counts) > 1, rejected)


def safe_qty(value: object) -> int:
    return _safe_int(value, 1)
