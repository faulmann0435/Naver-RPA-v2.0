"""Reusable Streamlit form for ONE dictionary entry: a plain sentence plus "which numbers scale" ticks.

Streamlit forgets the state of widgets that are not rendered, so every run mirrors the values into
`st.session_state["pending"][prefix]` and a re-rendered form starts from there.
"""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import replace

import streamlit as st

from core.template_builder import (
    build_template,
    guess_scale,
    numbers_of,
    sentence_and_scale,
)
from ui.option_logic import FormValues, form_issues, result_lines

__all__ = [
    "FormValues", "clear_pending", "clear_pending_with_prefix", "init_state", "option_form",
    "pending_values", "state_key", "vendor_select",
]

PENDING = "pending"
MAX_NUMBER_COLUMNS = 6
ADVANCED_HELP = (
    "같은 이름으로 합산하면 여러 옵션의 무게를 더해 한 줄로 적습니다. "
    "예) ★제수용 피문어 1kg + 0.5kg + (자숙) → ★제수용 피문어 1.5kg (자숙)"
)


def state_key(prefix: str, name: str) -> str:
    """session_state key of one widget of the form `prefix`."""
    return f"{prefix}__{name}"


_key = state_key


# ---------------------------------------------------------------- pending store

def pending_values(prefix: str) -> FormValues | None:
    return st.session_state.get(PENDING, {}).get(prefix)


def _set_pending(prefix: str, values: FormValues) -> None:
    st.session_state[PENDING] = {**st.session_state.get(PENDING, {}), prefix: values}


def _drop_widget_keys(prefix: str) -> None:
    for key in [k for k in st.session_state if isinstance(k, str) and k.startswith(prefix + "__")]:
        del st.session_state[key]


def clear_pending(prefix: str) -> None:
    """Forget the edits of one entry (stored values and widget state)."""
    st.session_state[PENDING] = {k: v for k, v in st.session_state.get(PENDING, {}).items() if k != prefix}
    _drop_widget_keys(prefix)


def clear_pending_with_prefix(start: str) -> None:
    """Forget every entry whose prefix starts with `start`."""
    for prefix in [k for k in st.session_state.get(PENDING, {}) if k.startswith(start)]:
        clear_pending(prefix)


# ---------------------------------------------------------------- state

def init_state(prefix: str, initial: FormValues) -> FormValues:
    """Make sure the widget keys exist (from the pending edits, else `initial`); returns the source values."""
    base = pending_values(prefix) or initial
    state = st.session_state
    if _key(prefix, "sentence") not in state:
        sentence, scale = sentence_and_scale(base.template)
        state[_key(prefix, "sentence")] = sentence
        state[_key(prefix, "prev")] = (sentence, tuple(scale))
        for i, flag in enumerate(scale):
            state[_key(prefix, f"scale_{i}")] = flag
    defaults = {
        "vendor": base.vendor, "qty1_on": bool(base.qty1_template), "qty1": base.qty1_template,
        "group": base.sum_group, "weight": base.unit_weight_kg, "append": base.append_to_end,
        "done": not base.needs_review, "enabled": base.enabled,
    }
    for name, value in defaults.items():
        state.setdefault(_key(prefix, name), value)
    return base


def vendor_select(prefix: str, initial: FormValues, vendor_ids: Sequence[str], label: str = "발주양식") -> str:
    init_state(prefix, initial)
    key = _key(prefix, "vendor")
    current = st.session_state[key]
    options = list(vendor_ids) if current in vendor_ids else [*vendor_ids, current]
    return str(st.selectbox(label, options, key=key))


def _current_scale(prefix: str, count: int) -> list[bool]:
    return [bool(st.session_state.get(_key(prefix, f"scale_{i}"), False)) for i in range(count)]


def _sync_scale(prefix: str, sentence: str) -> None:
    """When the sentence was edited, carry the ticks over to the new numbers."""
    old_sentence, old_scale = st.session_state[_key(prefix, "prev")]
    if sentence == old_sentence:
        return
    new_scale = guess_scale(old_sentence, old_scale, sentence)
    for i in range(len(numbers_of(old_sentence)) + len(new_scale)):
        key = _key(prefix, f"scale_{i}")
        if i < len(new_scale):
            st.session_state[key] = new_scale[i]
        elif key in st.session_state:
            del st.session_state[key]


# ---------------------------------------------------------------- pieces of the form

def _sentence_inputs(prefix: str) -> tuple[str, list[bool]]:
    grouped = bool(str(st.session_state[_key(prefix, "group")]).strip())
    sentence = st.text_input("발주서에 적을 문장", key=_key(prefix, "sentence"), disabled=grouped)
    if grouped:
        st.caption("합산 이름이 있어서 이 문장은 쓰이지 않습니다. (아래 '여러 옵션 합치기' 참고)")
    _sync_scale(prefix, sentence)
    numbers = numbers_of(sentence)
    if numbers and not grouped:
        st.markdown("수량만큼 늘릴 숫자:")
        for start in range(0, len(numbers), MAX_NUMBER_COLUMNS):
            chunk = numbers[start:start + MAX_NUMBER_COLUMNS]
            for offset, (col, number) in enumerate(zip(st.columns(len(chunk)), chunk, strict=True)):
                col.checkbox(number, key=_key(prefix, f"scale_{start + offset}"))
    scale = _current_scale(prefix, len(numbers))
    st.session_state[_key(prefix, "prev")] = (sentence, tuple(scale))
    return sentence, scale


def _qty1_input(prefix: str) -> str:
    if not st.toggle("1개일 때는 다르게 쓰기", key=_key(prefix, "qty1_on")):
        return ""
    return str(st.text_input("1개일 때 문장", key=_key(prefix, "qty1")))


def _advanced_inputs(prefix: str) -> tuple[str, float | None, bool]:
    state = st.session_state
    opened = bool(str(state[_key(prefix, "group")]).strip() or state[_key(prefix, "append")])
    with st.expander("여러 옵션 합치기 (고급)", expanded=opened):
        st.caption(ADVANCED_HELP)
        group = st.text_input("합산 이름", key=_key(prefix, "group"))
        weight = st.number_input(
            "1개당 무게(kg)", min_value=0.0, value=None, step=0.1, format="%g", key=_key(prefix, "weight")
        )
        append = st.checkbox("묶음 끝에 붙임", key=_key(prefix, "append"))
    return group.strip(), (None if weight is None else float(weight)), append


def _show_results(values: FormValues, vendor_ids: Sequence[str], qty_example: int) -> None:
    st.markdown("**결과:**")
    for qty, text in result_lines(values):
        st.text(f"{qty}개  →  {text}")
    for issue in form_issues(values, vendor_ids, qty_example):
        if issue.level == "error":
            st.error(issue.message)
        else:
            st.caption(f"⚠ {issue.message}")


def option_form(
    prefix: str,
    initial: FormValues,
    vendor_ids: Sequence[str],
    show_vendor: bool = True,
    show_flags: bool = False,
    qty_example: int = 0,
) -> FormValues:
    """Render the form of one entry and return its current values (also mirrored into `pending`)."""
    init_state(prefix, initial)
    if show_vendor:
        vendor_select(prefix, initial, vendor_ids)
    sentence, scale = _sentence_inputs(prefix)
    qty1 = _qty1_input(prefix)
    group, weight, append = _advanced_inputs(prefix)
    state = st.session_state
    values = FormValues(
        template=build_template(sentence, scale), qty1_template=qty1.strip(), vendor=str(state[_key(prefix, "vendor")]),
        sum_group=group, unit_weight_kg=weight, append_to_end=append,
        needs_review=not bool(state[_key(prefix, "done")]), enabled=bool(state[_key(prefix, "enabled")]),
    )
    _show_results(values, vendor_ids, qty_example)
    if show_flags:
        cols = st.columns(2)
        done = cols[0].checkbox("확인 완료", key=_key(prefix, "done"))
        enabled = cols[1].checkbox("사용", key=_key(prefix, "enabled"))
        values = replace(values, needs_review=not done, enabled=enabled)
    _set_pending(prefix, values)
    return values
