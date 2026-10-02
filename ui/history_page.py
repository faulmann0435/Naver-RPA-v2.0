"""🕓 변경 이력: revisions of one data file, the difference to the current file, and revert."""
from __future__ import annotations

import requests
import streamlit as st

from store.base import ConflictError, DataStore, Revision, StoreError
from store.dictionary_repo import load_settings
from store.rules_repo import OPTION_RULES_FILE, PRODUCT_ROUTE_FILE
from ui import context, dictionary_page, rules_state
from ui.history_logic import (
    TARGETS,
    diff_against_current,
    file_label,
    history_table,
    old_content,
    revert_errors,
    revert_file,
    revision_label,
)

HISTORY_LIMIT = 30
S_FLASH, S_CONFLICT = "history_flash", "history_conflict"


def _other_rules(store: DataStore, path: str) -> str:
    other = OPTION_RULES_FILE if path == PRODUCT_ROUTE_FILE else PRODUCT_ROUTE_FILE
    snapshot = store.read_text(other)
    return snapshot.content if snapshot else ""


def _forget_loaded_pages() -> None:
    st.session_state.pop(dictionary_page.S_BASE, None)
    rules_state.forget_loaded()


def _revert(store: DataStore, path: str, rev: Revision, sha: str | None) -> None:
    try:
        revert_file(store, path, rev, sha, context.current_author())
    except ConflictError:
        st.session_state[S_CONFLICT] = (path, old_content(store, path, rev))
        return
    except (StoreError, requests.RequestException) as e:
        st.error(f"되돌리지 못했습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    context.clear_data_caches()
    _forget_loaded_pages()
    st.session_state[S_FLASH] = f"되돌렸습니다: {file_label(path)} → {revision_label(rev)}"
    st.rerun()


def _conflict_box(path: str) -> None:
    saved = st.session_state.get(S_CONFLICT)
    if saved is None or saved[0] != path:
        return
    st.error("다른 사람이 먼저 저장했습니다. 화면을 새로 고친 뒤 다시 시도하세요.")
    st.download_button("되돌리려던 버전 받기", data=saved[1].encode("utf-8"), file_name=f"되돌리기_{path}",
                       key="history_mine")


def render() -> None:
    st.subheader("🕓 변경 이력")
    store = context.get_store()
    if store is None:
        st.info("데이터 저장소가 설정되지 않아 이력을 볼 수 없습니다. 설정(secrets)을 확인하세요.")
        return
    config = context.load_config_or_error()
    if config is None:
        return
    flash = st.session_state.pop(S_FLASH, None)
    if flash:
        st.success(flash)
    label = st.selectbox("대상", list(TARGETS), key="history_target")
    path = TARGETS[label]
    try:
        revisions = store.history(path, limit=HISTORY_LIMIT)
        current = store.read_text(path)
    except (StoreError, requests.RequestException) as e:
        st.error(f"이력을 불러올 수 없습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    if not revisions or current is None:
        st.info("이력이 없습니다.")
        return
    st.dataframe(history_table(revisions), hide_index=True, width="stretch")
    pick = st.selectbox("비교할 버전", range(len(revisions)), format_func=lambda i: revision_label(revisions[i]),
                        key=f"history_pick_{path}")
    rev = revisions[int(pick)]
    try:
        old = old_content(store, path, rev)
        settings = load_settings(store)
    except (StoreError, requests.RequestException, ValueError) as e:
        st.error(f"선택한 버전을 불러올 수 없습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    st.markdown("**선택한 버전과 현재 내용의 차이**")
    try:
        diff = diff_against_current(path, old, current.content, settings)
    except (ValueError, KeyError) as e:
        st.error(f"차이를 계산할 수 없습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    if diff.empty:
        st.success("현재 내용과 같습니다.")
    else:
        if not diff.table.empty:
            st.dataframe(diff.table, hide_index=True, width="stretch")
        if diff.note:
            st.info(diff.note)
    errors = revert_errors(path, old, context.vendor_ids(config), settings, _other_rules(store, path), context.load_raw_layout())
    for line in errors:
        st.error(f"이 버전으로는 되돌릴 수 없습니다: {line}")
    st.caption("되돌려도 이력은 지워지지 않고, 옛 내용이 '새 버전'으로 하나 더 쌓입니다. 언제든 다시 되돌릴 수 있습니다.")
    confirmed = st.checkbox("현재 내용을 이 버전으로 바꿉니다", key=f"history_confirm_{path}")
    if st.button("이 버전으로 되돌리기", disabled=bool(errors) or not confirmed, key="history_revert"):
        _revert(store, path, rev, current.sha)
    _conflict_box(path)

