"""Excel import of the item dictionary (inside the grid tab of 품목 관리)."""
from __future__ import annotations

from collections.abc import Callable

import pandas as pd
import requests
import streamlit as st

from core.dictionary import DictionarySettings
from store.base import ConflictError, StoreError
from store.csv_codec import to_csv_text
from store.dictionary_repo import now_kst_iso, save_dictionary
from ui import context
from ui.excel_import_logic import (
    ImportFormatError,
    ImportPlan,
    import_message,
    judge_import,
    plan_import,
    read_import_frame,
)

S_UPLOAD_N, S_CONFLICT = "dict_import_n", "dict_import_conflict"


def _apply(plan: ImportPlan, sha: str | None, reload: Callable[[], None]) -> None:
    message = import_message(plan)
    try:
        save_dictionary(context.get_store(), plan.frame, sha, context.current_author(), message)  # type: ignore[arg-type]
    except ConflictError:
        st.session_state[S_CONFLICT] = to_csv_text(plan.frame)
        return
    except (StoreError, requests.RequestException) as e:
        st.error(f"저장하지 못했습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    context.clear_data_caches()
    reload()
    st.session_state[S_UPLOAD_N] = st.session_state.get(S_UPLOAD_N, 0) + 1
    st.session_state["dict_flash"] = f"가져왔습니다. {message}"
    st.rerun()


def render_import(base: pd.DataFrame, sha: str | None, ids: list[str], reload: Callable[[], None]) -> None:
    st.markdown("**엑셀 가져오기**")
    st.caption(
        "'엑셀로 내보내기'로 받은 파일을 고쳐서 올리세요. 열 이름과 개수는 바꾸면 안 됩니다. "
        "표에서 고친 뒤 저장하지 않은 내용은 가져오기 후 사라집니다."
    )
    upload = st.file_uploader("품목 사전 엑셀 (.xlsx)", type=["xlsx"], key=f"dict_import_file_{st.session_state.get(S_UPLOAD_N, 0)}")
    if upload is None:
        return
    try:
        incoming = read_import_frame(upload.getvalue())
    except ImportFormatError as e:
        st.error(str(e))
        return
    state = context.load_dictionary_state_or_error()
    settings: DictionarySettings = state[1] if state else DictionarySettings()
    delete_missing = st.checkbox("엑셀에 없는 항목은 사전에서 삭제", value=False, key="dict_import_delete")
    plan = plan_import(base, incoming, settings, context.current_user(), now_kst_iso(), delete_missing)
    judged = judge_import(plan, ids, settings)
    st.write(f"추가 {plan.added} / 수정 {plan.modified} / 변경 없음 {plan.unchanged}"
             + (f" / 삭제 {plan.deleted}" if delete_missing else ""))
    if not plan.changes.empty:
        st.dataframe(plan.changes, hide_index=True, width="stretch")
    for line in judged.errors:
        st.error(line)
    if judged.warnings:
        with st.expander(f"경고 {len(judged.warnings)}건"):
            for line in judged.warnings:
                st.warning(line)
    if judged.reference:
        with st.expander(f"참고: 바뀌지 않은 항목의 문제 {len(judged.reference)}건 (가져오기를 막지 않음)"):
            for line in judged.reference:
                st.caption(line)
    nothing = not (plan.added or plan.modified or plan.deleted)
    if st.button("가져오기 적용", type="primary", disabled=nothing or bool(judged.errors), key="dict_import_apply"):
        _apply(plan, sha, reload)
    mine = st.session_state.get(S_CONFLICT)
    if mine is not None:
        st.error("다른 사람이 먼저 저장했습니다. 새로 불러온 뒤 다시 가져오세요.")
        st.download_button("가져오려던 결과 받기", data=mine.encode("utf-8"), file_name="dictionary_가져오기결과.csv",
                           mime="text/csv", key="dict_import_mine")
