import pytest

from core.template import (
    TemplateError,
    format_number,
    has_qty_placeholder,
    render,
    validate_template,
)


def test_multiply():
    assert render("마리 {수량*8}마리", 2) == "마리 16마리"


def test_decimal_multiply():
    assert render("{수량*0.5}", 2) == "1"
    assert render("{수량*0.5}", 3) == "1.5"


def test_spaces_inside_braces():
    assert render("{ 수량 }개 {   수량 * 8 }", 2) == "2개 16"


def test_plain_text_unchanged():
    assert render("깐멍게 500g", 3) == "깐멍게 500g"
    assert not has_qty_placeholder("깐멍게 500g")
    assert has_qty_placeholder("x{ 수량*2 }")


@pytest.mark.parametrize("bad", ["{수량", "{foo}", "{수량*}", "}", "{}", "{수량*8*2}", "a {수량}}"])
def test_invalid_templates(bad):
    assert validate_template(bad)
    with pytest.raises(TemplateError):
        render(bad, 1)


def test_valid_template_has_no_errors():
    assert validate_template("a {수량} b {수량*1.5}") == []


@pytest.mark.parametrize(
    ("value", "text"),
    [(1.0, "1"), (0.5, "0.5"), (1.25, "1.25"), (16.0, "16"), (0.3333333, "0.333"), (2.5, "2.5")],
)
def test_format_number(value, text):
    assert format_number(value) == text
