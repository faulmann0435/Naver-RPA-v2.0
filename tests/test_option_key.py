import math

from core.option_key import make_option_key, normalize_product_no

IGNORED = ["수령일 선택 (도착시간 지정불가)"]


def test_ignored_group_dropped_and_emoji_removed():
    raw = "수령일 선택 (도착시간 지정불가): 2월 5일 수령 / 가리비 선택: ⭐️국산참가리비 (대~특대) 1kg"
    assert make_option_key(raw, IGNORED) == "가리비 선택: 국산참가리비 (대~특대) 1kg"


def test_slash_inside_value_does_not_split():
    raw = "❤️문어 자숙으로 발송 무료이벤트: 한마리 통으로 찜해서 발송/ 찜비용 무료이벤트"
    assert make_option_key(raw) == "문어 자숙으로 발송 무료이벤트: 한마리 통으로 찜해서 발송/ 찜비용 무료이벤트"


def test_star_wrapped_value():
    raw = "속초 깔끔코다리 선택: ⭐속초 프리미엄코다리 8마리+소스 SET⭐"
    assert make_option_key(raw) == "속초 깔끔코다리 선택: 속초 프리미엄코다리 8마리+소스 SET"


def test_two_segments_kept():
    raw = "초코오징어 선택: ⭐한정수량 생물 오징어 / 마릿수 선택: 생물 오징어 800g-1kg내외 (2~3미)"
    expected = "초코오징어 선택: 한정수량 생물 오징어 / 마릿수 선택: 생물 오징어 800g-1kg내외 (2~3미)"
    assert make_option_key(raw) == expected


def test_hanja_kept_and_spaces_collapsed():
    assert make_option_key("크기 선택:   大   사이즈 ") == "크기 선택: 大 사이즈"


def test_missing_values():
    assert make_option_key(None) == ""
    assert make_option_key(math.nan) == ""


def test_slash_in_parentheses_not_split():
    raw = "수령일 선택 (도착시간 지정불가): 2/5 / 상품 선택: 광어(자숙/순차출고)"
    assert make_option_key(raw, IGNORED) == "상품 선택: 광어(자숙/순차출고)"


def test_product_no_variants():
    assert normalize_product_no(5568579375) == "5568579375"
    assert normalize_product_no(5568579375.0) == "5568579375"
    assert normalize_product_no(" 5568579375.0 ") == "5568579375"
    assert normalize_product_no("5568579375") == "5568579375"
    assert normalize_product_no(None) == ""
    assert normalize_product_no(math.nan) == ""
