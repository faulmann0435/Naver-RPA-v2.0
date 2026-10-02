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


def test_strip_invisible_removes_leftovers_but_keeps_visible_symbols():
    from core.template import strip_invisible

    assert strip_invisible("\ufe0f속초 오징어젓갈 220g x1") == "속초 오징어젓갈 220g x1"
    assert strip_invisible("\ufeff국산\u200d참가리비") == "국산참가리비"
    assert strip_invisible("★제수용 피문어") == "★제수용 피문어"  # intentional symbols stay


def test_render_entry_output_has_no_invisible_characters():
    from core.dictionary import DictionaryEntry, render_entry

    entry = DictionaryEntry(product_no="1", option_key="a", vendor_id="v",
                            display_template="\ufe0f속초 오징어젓갈 220g x{수량}")
    assert render_entry(entry, 2) == "속초 오징어젓갈 220g x2"


def test_clean_display_text_removes_emoji_keeps_stars_and_placeholders():
    from core.template import clean_display_text

    assert clean_display_text("🔥라면용 홍게 3kg") == "라면용 홍게 3kg"
    assert clean_display_text("함께즐기는 ❤️ 새콤달콤 명태회무침 220g x{수량}") == "함께즐기는 새콤달콤 명태회무침 220g x{수량}"
    assert clean_display_text("★제수용 피문어") == "★제수용 피문어"
    assert clean_display_text("\ufeff 🥇 백명란젓 {수량*2}개 ") == "백명란젓 {수량*2}개"
    assert clean_display_text("손질자숙피문어 통다리모듬 특大 {수량}개") == "손질자숙피문어 통다리모듬 특大 {수량}개"
