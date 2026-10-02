"""Seeding logic: probe the option rules, derive templates from hand text, build the CSV (synthetic data only)."""
from datetime import datetime, timezone

import pandas as pd

from core.dictionary import DictionarySettings
from core.seed import (
    HandEvidence,
    Sample,
    apply_evidence,
    derive_template,
    entry_display,
    expected_display,
    probe_entry,
    probe_row,
    probe_samples,
    strip_ignored_groups,
    verify_probe,
)
from store.csv_codec import from_csv_text, to_csv_text
from store.rules_repo import DICTIONARY_COLUMNS
from tools.seed_build import (
    MUNGE_VENDOR,
    SOURCE_AUTO,
    SOURCE_PAST_PO,
    build_seed,
    dictionary_frame,
)

RULE_COLUMNS = ["Order", "ApplyTo", "TargetKeyword", "ActionType", "Parameter", "Description"]


def rules(*rows: tuple[str, str, str]) -> pd.DataFrame:
    """OptionRules shaped like normalize_config output: rows are (ActionType, Parameter, TargetKeyword)."""
    data = [[i, "ALL", target, action, param, ""] for i, (action, param, target) in enumerate(rows, start=1)]
    return pd.DataFrame(data, columns=RULE_COLUMNS)


def config_for(option_rules: pd.DataFrame, keyword: str = "코다리", vendor: str = "속초 발주양식") -> dict:
    route = pd.DataFrame({"Priority": [1], "Keyword": [keyword], "TargetVendorID": [vendor]})
    return {"ProductRoute": route, "OptionRules": option_rules}


def test_probe_calc_unit_gives_multiplier_placeholder():
    probe = probe_row("코다리", "프리미엄 10마리", "V", rules(("CALC_UNIT", "마리", "ALL")))
    assert not probe.failed
    assert probe.template == "프리미엄 {수량*10}마리"
    entry = probe_entry(probe, "1", "k", "V")
    assert verify_probe(entry, probe)
    assert entry_display(entry, 3) == "프리미엄 30마리"


def test_probe_format_qty_gives_plain_qty_placeholder():
    probe = probe_row("코다리", "순대 1팩", "V", rules(("FORMAT_QTY", "x{qty}", "ALL")))
    assert not probe.failed
    assert probe.template == "순대 1팩 x{수량}"
    assert verify_probe(probe_entry(probe, "1", "k", "V"), probe)


def test_probe_plain_case_has_no_placeholder():
    probe = probe_row("코다리", "순대 500g 세트", "V", rules(("REMOVE_TEXT", "세트", "ALL")))
    assert not probe.failed
    assert probe.template == "순대 500g"
    entry = probe_entry(probe, "1", "k", "V")
    assert verify_probe(entry, probe)
    assert entry_display(entry, 2) == "순대 500g (x2)"


def test_probe_weight_gives_sum_group_and_unit_weight():
    probe = probe_row("장어", "손질장어 500g", "V", rules(("CONVERT_WEIGHT", "", "ALL")))
    assert not probe.failed
    assert probe.sum_group == "손질장어"
    assert probe.unit_weight_kg == 0.5
    entry = probe_entry(probe, "1", "k", "V")
    assert verify_probe(entry, probe)
    assert entry_display(entry, 3) == "손질장어 1.5kg"
    assert expected_display(probe.samples, 3) == "손질장어 1.5kg"


def test_probe_failure_falls_back_to_text_without_placeholder():
    samples = (Sample("순대 5개", 0.0, False), Sample("순대 7개", 0.0, True), Sample("순대 9개", 0.0, True))
    probe = probe_samples(samples)
    assert probe.failed
    assert probe.template == "순대 5개"


def test_strip_ignored_groups_keeps_other_text_untouched():
    text = "수령일 선택 (도착시간 지정불가): 2월 7일 / 가리비 선택: ⭐1kg"
    assert strip_ignored_groups(text, ("수령일 선택 (도착시간 지정불가)",)) == "가리비 선택: ⭐1kg"


def test_derive_template_from_hand_text():
    template = derive_template("16마리, 소스 2개", 2, "속초 프리미엄코다리 8마리+소스 SET")
    assert template == "{수량*8}마리, 소스 {수량}개"


def test_derive_template_keeps_unrelated_numbers_literal():
    assert derive_template("1kg (13미 내외)", 1, "1kg (13미 내외)") == "1kg (13미 내외)"
    assert derive_template("양미리 2두름", 2, "양미리 1두름") == "양미리 {수량}두름"
    assert derive_template("{수량}", 1, "x") is None


def _base_entry():
    probe = probe_row("코다리", "프리미엄 10마리", "V", rules(("CALC_UNIT", "마리", "ALL")))
    return probe_entry(probe, "1", "k", "V")


def test_evidence_prefers_probe_template_when_it_reproduces_hand_text():
    result = apply_evidence(_base_entry(), [HandEvidence("프리미엄 10마리", 1, "프리미엄 10마리")])
    assert result.used
    assert result.entry.display_template == "프리미엄 {수량*10}마리"


def test_evidence_conflict_takes_most_frequent_template():
    base = _base_entry()
    evidence = [
        HandEvidence("A 3개", 1, "x"),
        HandEvidence("B 3개", 1, "x"),
        HandEvidence("B 3개", 1, "x"),
    ]
    result = apply_evidence(base, evidence)
    assert result.used
    assert result.conflict
    assert result.entry.display_template == "B 3개"


def test_evidence_that_cannot_be_rendered_is_rejected():
    result = apply_evidence(_base_entry(), [HandEvidence("그냥 문구", 2, "x")])
    assert not result.used
    assert result.rejected == 1
    assert result.entry.display_template == "프리미엄 {수량*10}마리"


def _orders() -> pd.DataFrame:
    return pd.DataFrame({
        "상품번호": ["1", "2", "3", "4"],
        "상품명": ["코다리 A", "코다리 멍게", "기타", "코다리 A"],
        "옵션정보": ["프리미엄 10마리", "깐멍게 1개", "x", "수령일 선택 (도착시간 지정불가): 2월 7일 / 프리미엄 10마리"],
        "수량": [1, 1, 1, 1],
    })


def test_build_seed_skips_munge_and_unclassified_and_counts_them():
    config = config_for(rules(("CALC_UNIT", "마리", "ALL")), keyword="코다리", vendor=MUNGE_VENDOR)
    result = build_seed(_orders(), config, {}, DictionarySettings())
    assert result.skipped == {"멍게(사용자 판단 대기)": 1, "Unclassified": 1}
    assert [i.entry.product_no for i in result.items] == ["1", "4"]
    assert all(i.source == SOURCE_AUTO for i in result.items)


def test_build_seed_uses_past_po_evidence_and_ignores_date_group():
    config = config_for(rules(("CALC_UNIT", "마리", "ALL")))
    evidence = {("4", "프리미엄 10마리"): [HandEvidence("프리미엄 20마리", 2, "프리미엄 10마리")]}
    result = build_seed(_orders(), config, evidence, DictionarySettings())
    by_no = {i.entry.product_no: i for i in result.items}
    assert by_no["4"].source == SOURCE_PAST_PO
    assert by_no["4"].entry.display_template == "프리미엄 {수량*10}마리"
    assert "2월" not in by_no["4"].entry.display_template
    assert not result.probe_failures and not result.verify_failures


def test_csv_has_exactly_dictionary_columns_and_round_trips():
    config = config_for(rules(("CALC_UNIT", "마리", "ALL")))
    result = build_seed(_orders(), config, {}, DictionarySettings())
    frame = dictionary_frame(result.items, datetime(2026, 1, 1, tzinfo=timezone.utc))
    assert list(frame.columns) == DICTIONARY_COLUMNS
    parsed = from_csv_text(to_csv_text(frame), as_text=True)
    assert list(parsed.columns) == DICTIONARY_COLUMNS
    assert set(parsed["updated_by"]) == {"seed"}
    assert set(parsed["needs_review"]) == {"1"}
    assert list(parsed["vendor_id"]) == sorted(parsed["vendor_id"])


def test_probe_ignores_stray_qty1_stamp_and_uses_qty_2_3():
    samples = (
        Sample("코다리 10마리+소스 (x1)", 0.0, True),
        Sample("코다리 20마리+소스", 0.0, True),
        Sample("코다리 30마리+소스", 0.0, True),
    )
    probe = probe_samples(samples)
    assert not probe.failed
    assert probe.q1_differs
    assert probe.template == "코다리 {수량*10}마리+소스"
    assert verify_probe(probe_entry(probe, "1", "k", "V"), probe)
