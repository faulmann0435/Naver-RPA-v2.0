"""고급 설정 > 설정 tab: item separator and ignored option groups (settings.json)."""
from __future__ import annotations

import pandas as pd
import requests
import streamlit as st

from store.base import ConflictError, DataStore, StoreError
from store.rules_repo import SETTINGS_FILE
from ui import context, rules_state
from ui.rules_extra_logic import (
    SEPARATOR_CHOICES,
    build_settings_text,
    clean_groups,
    parse_settings,
    settings_values,
    validate_settings,
)

CUSTOM = "직접 입력"
GROUP_COLUMN = "옵션 항목 이름"
S_SET_CONFLICT = "rules_conflict_settings"


def _choose_separator(current: str, sha: str | None) -> str:
    choices = [*SEPARATOR_CHOICES, CUSTOM]
    index = choices.index(current) if current in SEPARATOR_CHOICES else len(choices) - 1
    chosen = st.selectbox(
        "품목 구분 기호", choices, index=index, format_func=lambda v: v if v == CUSTOM else f'"{v}"',
        key=f"rules_w_sep_{sha}",
    )
    st.caption("한 상자에 여러 품목이 들어갈 때 품목 사이에 넣는 기호입니다.")
    if chosen != CUSTOM:
        return str(chosen)
    return st.text_input("직접 입력할 구분 기호", value=current, key=f"rules_w_sep_custom_{sha}")


def _save(store: DataStore, text: str, sha: str | None, message: str) -> None:
    try:
        store.write_text(SETTINGS_FILE, text, sha, message, context.current_author())
    except ConflictError:
        st.session_state[S_SET_CONFLICT] = text
        return
    except (StoreError, requests.RequestException) as e:
        st.error(f"저장하지 못했습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    context.clear_data_caches()
    rules_state.load_all(store)
    st.session_state[rules_state.S_FLASH] = f"저장했습니다. {message}"
    st.rerun()


def render_settings_tab(store: DataStore) -> None:
    text, sha = st.session_state[rules_state.S_SETTINGS]
    try:
        existing = parse_settings(text)
    except ValueError as e:
        st.error(f"settings.json을 읽을 수 없습니다. ({e})")
        return
    separator, groups = settings_values(existing)
    st.caption(f"설정 파일 버전 {(sha or '없음')[:7]} 기준으로 편집 중")
    new_separator = _choose_separator(separator, sha)
    st.caption("비교 제외 옵션 항목: 옵션 이름이 달라도 같은 옵션으로 보기 위해 무시하는 옵션 항목입니다 (예: 수령일 선택).")
    edited = st.data_editor(
        pd.DataFrame({GROUP_COLUMN: groups}, dtype=str), num_rows="dynamic", hide_index=True, width="stretch",
        key=f"rules_w_groups_{sha}",
    )
    new_groups = clean_groups(edited[GROUP_COLUMN].tolist())
    errors = validate_settings(new_separator)
    changed = (new_separator, new_groups) != (separator, groups)
    for line in errors:
        st.error(line)
    if st.button("저장", type="primary", disabled=not changed or bool(errors), key="rules_w_save_settings"):
        _save(store, build_settings_text(existing, new_separator, new_groups), sha, "설정 수정: 품목 구분 기호/비교 제외 옵션 항목")
    conflict = st.session_state.get(S_SET_CONFLICT)
    if conflict is not None:
        st.error("다른 사람이 먼저 저장했습니다. '새로 불러오기' 후 다시 수정하세요.")
        st.download_button("내 수정본 받기", data=conflict.encode("utf-8"), file_name="settings_내수정본.json",
                           mime="application/json", key="rules_w_mine_settings")

