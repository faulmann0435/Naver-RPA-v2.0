"""Markdown report of the seeding run. Product-level data only: no customer names, phones or addresses."""
from __future__ import annotations

from dataclasses import dataclass

from tools.seed_build import KeyNote, SeedResult
from tools.seed_data import OrdersLoad
from tools.seed_eval import BundleResult

MAX_TEXT = 220


@dataclass(frozen=True)
class PoEval:
    label: str
    results: list[BundleResult]
    po_rows: int
    unmatched_rows: int
    ambiguous_rows: int
    no_hand_bundles: int


@dataclass(frozen=True)
class Totals:
    bundles: int
    exact_a: int
    exact_b: int
    norm_a: int
    norm_b: int


def totals(evals: list[PoEval]) -> Totals:
    rows = [r for e in evals for r in e.results]
    return Totals(
        len(rows),
        sum(r.exact_a for r in rows),
        sum(r.exact_b for r in rows),
        sum(r.norm_a for r in rows),
        sum(r.norm_b for r in rows),
    )


def _clip(text: str) -> str:
    flat = " ".join(text.split())
    return flat if len(flat) <= MAX_TEXT else flat[:MAX_TEXT] + "..."


def _cell(text: str) -> str:
    return _clip(text).replace("|", "\\|")


def _notes_section(title: str, notes: list[KeyNote]) -> list[str]:
    lines = [f"### {title} ({len(notes)})", ""]
    if not notes:
        return [*lines, "(none)", ""]
    lines += ["| 상품번호 | 상품명 | 옵션정보 |", "|---|---|---|"]
    lines += [f"| {n.product_no} | {_cell(n.product_name)} | {_cell(n.option_raw)} |" for n in notes]
    return [*lines, ""]


def _eval_table(evals: list[PoEval], total: Totals) -> list[str]:
    lines = [
        "| PO file | bundles | exact A | exact B | normalized A | normalized B | PO rows | unmatched joins | ambiguous joins | no hand text |",
        "|---|---|---|---|---|---|---|---|---|---|",
    ]
    for e in evals:
        t = totals([e])
        lines.append(
            f"| {e.label} | {t.bundles} | {t.exact_a} | {t.exact_b} | {t.norm_a} | {t.norm_b} "
            f"| {e.po_rows} | {e.unmatched_rows} | {e.ambiguous_rows} | {e.no_hand_bundles} |"
        )
    lines.append(
        f"| **Total** | {total.bundles} | {total.exact_a} | {total.exact_b} | {total.norm_a} | {total.norm_b} "
        f"| {sum(e.po_rows for e in evals)} | {sum(e.unmatched_rows for e in evals)} "
        f"| {sum(e.ambiguous_rows for e in evals)} | {sum(e.no_hand_bundles for e in evals)} |"
    )
    return lines


def _mismatches(evals: list[PoEval]) -> list[str]:
    lines = ["| PO file | # | hand text | B output | A output |", "|---|---|---|---|---|"]
    count = 0
    for e in evals:
        for r in e.results:
            if r.norm_b:
                continue
            count += 1
            b_out = _cell(" | ".join(r.b_texts)) if r.b_texts else "(no output)"
            a_out = _cell(" | ".join(r.a_texts)) if r.a_texts else "(no output)"
            lines.append(f"| {e.label} | #{r.index} | {_cell(r.hand)} | {b_out} | {a_out} |")
    return lines if count else ["(none)"]


def render_report(
    seed: SeedResult,
    load: OrdersLoad,
    evals: list[PoEval],
    pairs_found: int,
    pairs_skipped: int,
    password_used: bool,
) -> str:
    total = totals(evals)
    sources = [i.source for i in seed.items]
    lines = [
        "# Draft dictionary seed report",
        "",
        (
            "> Caution: the dictionary was partly built from these same past POs, so the B (with dictionary) "
            "numbers are optimistic. Treat them as an upper bound, not as accuracy on new orders."
        ),
        "",
        "## Data",
        "",
        f"- Order files read: {load.files_read} (skipped: {load.files_skipped}; password given: {'yes' if password_used else 'no'})",
        f"- Order rows: {load.rows_total} (duplicates by 상품주문번호 removed: {load.rows_duplicate})",
        f"- PO pairs found: {pairs_found} (evaluated: {pairs_found - pairs_skipped}, unreadable or encrypted: {pairs_skipped})",
        "",
        "## Dictionary",
        "",
        f"- Entries: {len(seed.items)} (past_po: {sources.count('past_po')}, auto_rule: {sources.count('auto_rule')})",
        f"- Skipped keys: {sum(seed.skipped.values())} " + ", ".join(f"[{k}: {v}]" for k, v in seed.skipped.items()),
        f"- Probe failures: {len(seed.probe_failures)}; verification failures: {len(seed.verify_failures)}",
        f"- Entries whose qty-1 engine output carries a stray 'x1' stamp (template built from qty 2-3): {seed.q1_differs}",
        f"- Conflicting hand-text evidence: {len(seed.conflicts)}; rejected evidence items: {seed.rejected_evidence}",
        "",
        *_notes_section("Probe failures (template = text for qty 1, no placeholder)", seed.probe_failures),
        *_notes_section("Verification failures (entry does not reproduce the classic output for qty 1-3)", seed.verify_failures),
        *_notes_section("Conflicting evidence (most frequent template used)", seed.conflicts),
        "## Evaluation: A = without dictionary, B = with seeded dictionary",
        "",
        *_eval_table(evals, total),
        "",
        "### Mismatches for B (normalized), identified by bundle index",
        "",
        "append_to_end is not inferred automatically; bundles that remain mismatched below are candidates for it.",
        "",
        *_mismatches(evals),
        "",
    ]
    return "\n".join(lines)
