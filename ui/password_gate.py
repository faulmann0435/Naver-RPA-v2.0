"""Shared-password screen for the test deployment.

The password lives only in Streamlit Secrets (`[app_gate]` / `password`). When it is absent
or empty there is no gate at all (local dev, and the production app with email-invite access).
The entered password is never stored, logged or echoed.
"""
from __future__ import annotations

import hmac
import time
from collections.abc import Mapping
from typing import Any

import streamlit as st

GATE_SECTION = "app_gate"
GATE_KEY = "password"
PASSED_KEY = "_gate_passed"
FAILURES_KEY = "_gate_failures"
LOCKED_UNTIL_KEY = "_gate_locked_until"
MAX_FAILURES = 5
LOCKOUT_SECONDS = 60
GATE_TITLE = "속초 발주 처리 시스템 (시험용)"


# ---------------------------------------------------------------- pure helpers

def gate_password(secrets_like: Any) -> str | None:
    """The configured password, or None when the gate is not configured (missing/empty)."""
    try:
        if GATE_SECTION not in secrets_like:
            return None
        section = secrets_like[GATE_SECTION]
        if not isinstance(section, Mapping) and not hasattr(section, "get"):
            return None
        value = section.get(GATE_KEY)
    except Exception:  # noqa: BLE001  # no secrets file / malformed section
        return None
    if value is None:
        return None
    text = str(value)
    return text if text.strip() else None


def password_matches(entered: str, expected: str) -> bool:
    """Constant-time comparison on UTF-8 bytes; surrounding whitespace of the input is ignored."""
    return hmac.compare_digest(entered.strip().encode("utf-8"), expected.encode("utf-8"))


def _configured_password() -> str | None:
    try:
        return gate_password(st.secrets)
    except Exception:  # noqa: BLE001  # st.secrets raises when there is no secrets file
        return None


def _now() -> float:
    return time.monotonic()


# ---------------------------------------------------------------- screen

def _remaining_lock_seconds() -> int:
    until = float(st.session_state.get(LOCKED_UNTIL_KEY, 0.0))
    return max(0, int(until - _now() + 0.999))


def _register_failure() -> None:
    failures = int(st.session_state.get(FAILURES_KEY, 0)) + 1
    if failures >= MAX_FAILURES:
        st.session_state[FAILURES_KEY] = 0
        st.session_state[LOCKED_UNTIL_KEY] = _now() + LOCKOUT_SECONDS
    else:
        st.session_state[FAILURES_KEY] = failures


def _try_login(entered: str, expected: str) -> None:
    if password_matches(entered, expected):
        st.session_state[PASSED_KEY] = True
        st.session_state[FAILURES_KEY] = 0
        st.rerun()
    _register_failure()
    wait = _remaining_lock_seconds()
    if wait:
        st.error(f"비밀번호를 {MAX_FAILURES}번 틀렸습니다. {wait}초 뒤에 다시 시도하세요.")
    else:
        st.error("비밀번호가 맞지 않습니다.")


def require_password() -> bool:
    """True when the app may render: no gate configured, or the password was already accepted."""
    expected = _configured_password()
    if expected is None or st.session_state.get(PASSED_KEY):
        return True
    st.title(GATE_TITLE)
    with st.form("password_gate"):
        entered = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("들어가기")
    if submitted:
        wait = _remaining_lock_seconds()
        if wait:
            st.error(f"시도 횟수를 넘었습니다. {wait}초 뒤에 다시 시도하세요.")
        else:
            _try_login(entered, expected)
    return False
