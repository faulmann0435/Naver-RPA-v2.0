"""Help texts of the 고급 설정 page: what every ActionType does (shown to non-developers)."""
from __future__ import annotations

ACTION_HELP_COLUMNS = ["ActionType", "하는 일", "Parameter 예시", "결과 예시"]
ACTION_HELP: list[tuple[str, str, str, str]] = [
    ("REMOVE_TEXT", "쉼표로 나눈 문구를 그대로 지움", "특가, 상품 선택: ", "[특가] 문어 → 문어"),
    ("REMOVE_REGEX", "패턴(정규식)에 맞는 부분을 지움", r"\(.*?\)", "문어 (생물) → 문어"),
    ("REPLACE_REGEX_SUB", "`패턴 /// 바꿀 말` 형식으로 바꿈", "무라벨 /// 쭈구리", "무라벨 식혜 → 쭈구리 식혜"),
    ("MASK_TEXT / UNMASK_TEXT", "특정 문구를 다른 규칙에서 보호했다가 다시 풂", "1.5L", "–"),
    ("CALC_UNIT", "숫자+단위의 숫자에 수량을 곱함 (범위 `~`,`-`와 `내외`는 제외)", "마리", "10마리(수량 2) → 20마리"),
    (
        "CONVERT_WEIGHT",
        "괄호 밖 무게를 kg으로 바꿔 수량만큼 곱하고, 문구에서는 지운 뒤 발주서에는 묶음 합계 무게를 붙임 (한 행에 한 번)",
        "kg", "장어 1kg(수량 2) → 장어 2kg",
    ),
    ("APPEND_QTY_UNIT", '끝에 "수량+단위"를 붙임', "팩", "생물오징어(수량 3) → 생물오징어 3팩"),
    ("GROUP_MULTIPLY", "상품명·옵션에 이 글자가 있으면 끝에 `(x수량)`을 붙임", "소스", "매운맛소스(수량 2) → 매운맛소스 (x2)"),
    ("FORMAT_QTY", "수량 표기를 서식대로 한 번 붙임 (`{qty}` 자리에 수량)", "x{qty}", "명이나물(수량 2) → 명이나물 x2"),
    ("APPEND_SUFFIX", "끝에 문구 추가", "(냉동)", "문어 → 문어 (냉동)"),
    ("PREPEND_TEXT", "앞에 문구 추가", "★", "문어 → ★ 문어"),
]
ACTION_HELP_NOTE = "수량 표시 규칙(CALC_UNIT, APPEND_QTY_UNIT, GROUP_MULTIPLY, FORMAT_QTY)은 한 행에 한 번만 적용됩니다."
