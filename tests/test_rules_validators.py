"""Rule-table validators: every rule, the shadowing example and the final guard."""
import pandas as pd

from core.config_loader import read_config_sheets
from store.rules_repo import sheets_to_rule_csvs
from store.rules_validators import (
    final_guard,
    validate_option_rules,
    validate_product_route,
)

IDS = ["속초 발주양식", "문어 발주 양식"]


def route(*rows: tuple[str, str, str]) -> pd.DataFrame:
    return pd.DataFrame([{"priority": p, "keyword": k, "vendor": v, "enabled": "1"} for p, k, v in rows])


def options(*rows: tuple[str, str, str, str, str]) -> pd.DataFrame:
    return pd.DataFrame(
        [{"order": o, "vendor": v, "target": t, "action": a, "param": p, "enabled": "1"} for o, v, t, a, p in rows]
    )


def messages(issues, level):
    return [i.message for i in issues if i.level == level]


DEFAULT_ROW = ("999", "DEFAULT", IDS[0])


def test_route_valid_has_no_issues():
    assert validate_product_route(route(("1", "문어", IDS[1]), DEFAULT_ROW), IDS) == []


def test_route_priority_vendor_and_empty_keyword_errors():
    issues = validate_product_route(route(("a", "문어", IDS[1]), ("2", "홍어", "없는양식"), ("3", " , ", IDS[0]), DEFAULT_ROW), IDS)
    errors = [(i.row, i.column) for i in issues if i.level == "error"]
    assert errors == [(0, "priority"), (1, "vendor"), (2, "keyword")]


def test_route_needs_exactly_one_default():
    none = validate_product_route(route(("1", "문어", IDS[1])), IDS)
    two = validate_product_route(route(DEFAULT_ROW, ("1000", "default", IDS[0])), IDS)
    assert any("DEFAULT" in m and "없습니다" in m for m in messages(none, "error"))
    assert any("2개" in m for m in messages(two, "error"))
    assert none[0].row is None


def test_route_disabled_default_does_not_count():
    frame = route(DEFAULT_ROW)
    frame["enabled"] = "0"
    assert any("DEFAULT" in m for m in messages(validate_product_route(frame, IDS), "error"))


def test_route_duplicate_keyword_warns():
    issues = validate_product_route(route(("1", "문어, 홍게", IDS[1]), ("2", "홍게", IDS[0]), DEFAULT_ROW), IDS)
    assert any("중복" in m and "홍게" in m for m in messages(issues, "warning"))


def test_route_shadowed_keyword_names_both():
    issues = validate_product_route(route(("1", "홍게", IDS[0]), ("2", "라면용홍게", IDS[1]), DEFAULT_ROW), IDS)
    shadow = [i for i in issues if i.level == "warning"]
    assert shadow and all("홍게" in i.message and "라면용홍게" in i.message for i in shadow)
    assert {i.row for i in shadow} == {0, 1} and not messages(issues, "error")


def test_route_same_priority_earlier_row_shadows_but_later_priority_does_not_shadow_earlier():
    same = validate_product_route(route(("2", "홍게", IDS[0]), ("2", "라면용홍게", IDS[1]), DEFAULT_ROW), IDS)
    reverse = validate_product_route(route(("1", "라면용홍게", IDS[1]), ("2", "홍게", IDS[0]), DEFAULT_ROW), IDS)
    assert messages(same, "warning") and not messages(reverse, "warning")


def test_options_errors():
    frame = options(
        ("x", "ALL", "ALL", "REMOVE_TEXT", "a"),
        ("2", "ALL", "ALL", "NOPE", "a"),
        ("3", "없는양식", "ALL", "REMOVE_TEXT", "a"),
        ("4", "ALL", "ALL", "REMOVE_REGEX", "("),
        ("5", "ALL", "ALL", "REPLACE_REGEX_SUB", "( /// x"),
        ("6", "ALL", "ALL", "REPLACE_REGEX_SUB", "(a || x"),
    )
    errors = [(i.row, i.column) for i in validate_option_rules(frame, IDS) if i.level == "error"]
    assert errors == [(0, "order"), (1, "action"), (2, "vendor"), (3, "param"), (4, "param"), (5, "param")]


def test_options_regex_replacement_part_is_not_compiled():
    assert not messages(validate_option_rules(options(("1", "ALL", "ALL", "REPLACE_REGEX_SUB", r"a /// ( x")), IDS), "error")


def test_options_empty_parameter_errors_only_for_listed_actions():
    listed = ["REMOVE_TEXT", "REMOVE_REGEX", "REPLACE_REGEX_SUB", "MASK_TEXT", "CALC_UNIT", "GROUP_MULTIPLY",
              "APPEND_SUFFIX", "PREPEND_TEXT"]
    for action in listed:
        assert messages(validate_option_rules(options(("1", "ALL", "ALL", action, "")), IDS), "error"), action
    for action in ["UNMASK_TEXT", "CONVERT_WEIGHT", "APPEND_QTY_UNIT", "FORMAT_QTY"]:
        assert not messages(validate_option_rules(options(("1", "ALL", "ALL", action, "")), IDS), "error"), action


def test_options_warnings():
    frame = options(
        ("1", "ALL", "ALL", "REPLACE_REGEX_SUB", "abc"),
        ("2", "ALL", "ALL", "FORMAT_QTY", "x"),
        ("3", "ALL", "문어, 오징어", "REMOVE_TEXT", "a"),
        ("4", "ALL", "", "REMOVE_TEXT", "a"),
    )
    issues = validate_option_rules(frame, IDS)
    assert not messages(issues, "error")
    assert [(i.row, i.column) for i in issues] == [(0, "param"), (1, "param"), (2, "target"), (3, "target")]


def test_options_disabled_rows_are_skipped():
    frame = options(("1", "ALL", "ALL", "SET_UNIT_FLAG", ""))
    frame["enabled"] = "0"
    assert validate_option_rules(frame, IDS) == []


def test_final_guard_accepts_real_rules_and_rejects_broken_ones():
    raw = read_config_sheets("config.xlsx", "1111")
    files = sheets_to_rule_csvs(raw, "t", "2026-01-01")
    route_csv, rules_csv = files["product_route.csv"], files["option_rules.csv"]
    assert final_guard(route_csv, rules_csv, raw["OutputLayout"]) == []
    header = rules_csv.splitlines()[0]
    broken = final_guard(route_csv, header + "\n", raw["OutputLayout"])
    assert broken and broken[0].level == "error" and broken[0].row is None
    no_column = final_guard(route_csv, "순서,x\n1,2\n", raw["OutputLayout"])
    assert no_column
