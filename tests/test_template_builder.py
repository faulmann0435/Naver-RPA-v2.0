"""core.template_builder: sentence <-> template with scaled-number flags."""
import random

import pytest

from core.template import render
from core.template_builder import (
    Piece,
    build_template,
    guess_scale,
    sentence_and_scale,
    split_pieces,
)


def test_split_pieces_alternates_and_handles_decimals():
    assert split_pieces("1.5L") == [Piece("1.5", True), Piece("L", False)]
    assert split_pieces("16-18미") == [Piece("16", True), Piece("-", False), Piece("18", True), Piece("미", False)]
    assert split_pieces("") == []
    assert split_pieces("문구") == [Piece("문구", False)]
    assert split_pieces("a3") == [Piece("a", False), Piece("3", True)]


@pytest.mark.parametrize(("sentence", "scale", "expected"), [
    ("속초 프리미엄코다리 8마리, 소스 1개", [True, True], "속초 프리미엄코다리 {수량*8}마리, 소스 {수량}개"),
    ("깐멍게 500gx1개", [False, True], "깐멍게 500gx{수량}개"),
    ("수입산참가리비(중 16-18미) 1kg", [False, False, True], "수입산참가리비(중 16-18미) {수량}kg"),
    ("문어 0.5kg", [True], "문어 {수량*0.5}kg"),
    ("문어 0.5kg", [], "문어 0.5kg"),
    ("문어 0.5kg", [False, True, True], "문어 0.5kg"),
    ("고정 문구", [True], "고정 문구"),
    ("2", [True], "{수량*2}"),
    ("1", [True], "{수량}"),
    ("", [], ""),
])
def test_build_template(sentence, scale, expected):
    assert build_template(sentence, scale) == expected


def test_build_template_renders_as_expected():
    template = build_template("코다리 8마리", [True])
    assert render(template, 3) == "코다리 24마리"


@pytest.mark.parametrize(("template", "sentence", "scale"), [
    ("속초 코다리 {수량*8}마리, 소스 {수량}개", "속초 코다리 8마리, 소스 1개", [True, True]),
    ("깐멍게 500gx{수량}개", "깐멍게 500gx1개", [False, True]),
    ("A {수량 * 8} B", "A 8 B", [True]),
    ("{ 수량 }개", "1개", [True]),
])
def test_sentence_and_scale(template, sentence, scale):
    assert sentence_and_scale(template) == (sentence, scale)


def test_invalid_template_returned_unchanged():
    for template in ("{수량*}", "{abc}", "a {수량", "}"):
        assert sentence_and_scale(template) == (template, [])


def test_ambiguous_template_falls_back_but_still_round_trips():
    for template in ("{수량}5", "5{수량}", "{수량}{수량}", "{수량*1}", "3.{수량}"):
        sentence, scale = sentence_and_scale(template)
        assert (sentence, scale) == (template, [])
        assert build_template(sentence, scale) == template


@pytest.mark.parametrize("template", [
    "", "고정 문구", "코다리 {수량}마리", "{수량*8}마리", "a 10 b {수량*0.5} c 2.5 d", "(중 16-18미) {수량}kg",
    "x{수량}y{수량*3}z7",
])
def test_round_trip_examples(template):
    assert build_template(*sentence_and_scale(template)) == template


def test_round_trip_random_templates():
    rng = random.Random(7)
    words = ["문어", " ", "-", "x", "(", ")", "마리", "kg", ", ", "g"]
    for _ in range(500):
        parts = []
        for _ in range(rng.randint(0, 8)):
            kind = rng.choice(["word", "num", "qty", "mul"])
            if kind == "word":
                parts.append(rng.choice(words))
            elif kind == "num":
                parts.append(rng.choice(["2", "10", "0.5", "16", "1"]))
            elif kind == "qty":
                parts.append("{수량}")
            else:
                parts.append("{수량*" + rng.choice(["8", "0.5", "12", "3"]) + "}")
        template = "".join(parts)
        assert build_template(*sentence_and_scale(template)) == template, template


def test_round_trip_sentence_and_scale_random():
    rng = random.Random(11)
    for _ in range(300):
        sentence = "".join(
            rng.choice(["문어 ", "8", "1", "0.5", "kg", "-", " ", "x"]) for _ in range(rng.randint(0, 9))
        )
        count = sum(p.is_number for p in split_pieces(sentence))
        scale = [rng.random() < 0.5 for _ in range(count)]
        template = build_template(sentence, scale)
        assert build_template(*sentence_and_scale(template)) == template


@pytest.mark.parametrize(("old", "flags", "new", "expected"), [
    ("코다리 8마리 소스 1개", [True, False], "코다리 9마리 소스 1개", [True, False]),   # same count: by position
    ("코다리 8마리 소스 1개", [True, True], "코다리 8마리", [True]),                    # fewer: by value
    ("코다리 8마리", [True], "코다리 8마리 소스 1개", [True, False]),                    # more: by value
    ("코다리 8마리", [False], "코다리 8마리 8개", [False, False]),
    ("a 1 b 2", [False, True], "a 2 b", [True]),
    ("", [], "문어 3개", [False]),
    ("문어 3개", [True], "문어", []),
    ("문어 3개", [], "문어 3개 4개", [False, False]),
])
def test_guess_scale(old, flags, new, expected):
    assert guess_scale(old, flags, new) == expected


def test_guess_scale_compares_values_not_text():
    assert guess_scale("a 8 b 1", [True, False], "a 8.0") == [True]
