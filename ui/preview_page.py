"""🧪 결과 확인: try a bundle of items and see the purchase-order text without uploading a file."""
from __future__ import annotations

import pandas as pd
import streamlit as st

from core.dictionary import DictionarySettings, ItemDictionary
from ui import context
from ui.dictionary_logic import dictionary_from_work
from ui.dictionary_page import S_WORK
from ui.order_page import S_DF
from ui.preview_logic import PREVIEW_COLUMNS, PreviewResult, clean_rows, preview_bundle

S_ROWS, S_VERSION, S_PICK = "preview_rows", "preview_version", "preview_pick"
PLACEHOLDER = "(선택하세요)"


def _empty_rows() -> pd.DataFrame:
    return pd.DataFrame({"상품번호": [""], "상품명": [""], "옵션정보": [""], "수량": [1]})


def _recent_choices() -> dict[str, dict]:
    df = st.session_state.get(S_DF)
    if df is None or not {"상품명", "옵션정보"} <= set(df.columns):
        return {}
    numbers = df["상품번호"] if "상품번호" in df.columns else pd.Series([""] * len(df), index=df.index)
    choices: dict[str, dict] = {}
    for no, name, option in zip(numbers, df["상품명"], df["옵션정보"]):
        row = {"상품번호": "" if pd.isna(no) else str(no), "상품명": str(name), "옵션정보": "" if pd.isna(option) else str(option), "수량": 1}
        label = f"{row['상품명']} | {row['옵션정보']}"
        if label in choices and choices[label]["상품번호"] != row["상품번호"]:
            label = f"{label} [{row['상품번호']}]"
        choices.setdefault(label, row)
    return choices


def _append_picked(choices: dict[str, dict]) -> None:
    label = st.session_state.get(S_PICK)
    if label in choices:
        rows: pd.DataFrame = st.session_state[S_ROWS]
        st.session_state[S_ROWS] = pd.concat([rows, pd.DataFrame([choices[label]])], ignore_index=True)
        st.session_state[S_VERSION] += 1
    st.session_state[S_PICK] = PLACEHOLDER


def _show_result(result: PreviewResult, rows: pd.DataFrame) -> None:
    if not result.vendors:
        st.info("입력한 품목이 없습니다.")
        return
    st.markdown("**발주서에 들어갈 품목**")
    st.dataframe(
        pd.DataFrame([{"발주양식": v.vendor if v.vendor != "Unclassified" else "미분류", "품목": v.text} for v in result.vendors]),
        hide_index=True, width="stretch",
    )
    st.markdown("**행별 처리 경로**")
    paths = rows.reset_index(drop=True).assign(경로=result.paths)[["상품명", "옵션정보", "수량", "경로"]]
    st.dataframe(paths, hide_index=True, width="stretch")
    for i, log in result.debug.items():
        with st.expander(f"자동 규칙 적용 과정: {rows.reset_index(drop=True).at[i, '상품명']}"):
            st.code("\n".join(log) or "(적용된 규칙 기록 없음)")


def render() -> None:
    st.subheader("🧪 결과 확인")
    st.caption("여러 품목을 한 상자로 묶었을 때 결과를 봅니다.")
    config = context.load_config_or_error()
    if config is None:
        return
    st.session_state.setdefault(S_ROWS, _empty_rows())
    st.session_state.setdefault(S_VERSION, 0)
    choices = _recent_choices()
    if choices:
        st.selectbox("최근 주문에서 고르기", [PLACEHOLDER, *choices], key=S_PICK,
                     on_change=_append_picked, args=(choices,))
    edited = st.data_editor(
        st.session_state[S_ROWS], key=f"preview_editor_{st.session_state[S_VERSION]}", num_rows="dynamic",
        hide_index=True, width="stretch",
        column_config={
            "수량": st.column_config.NumberColumn("수량", min_value=1, step=1, default=1),
        },
    )
    st.session_state[S_ROWS] = edited[PREVIEW_COLUMNS]
    work = st.session_state.get(S_WORK)
    use_work = st.checkbox("편집 중인(저장 안 한) 사전으로 미리보기", disabled=work is None, key="preview_use_work")
    if not st.button("결과 보기"):
        return
    state = context.load_dictionary_state_or_error()
    if state is None:
        return
    dictionary: ItemDictionary = state[0]
    settings: DictionarySettings = state[1]
    if use_work and work is not None:
        dictionary = dictionary_from_work(work, settings)
    try:
        result = preview_bundle(edited, config, dictionary, settings)
    except Exception as e:  # noqa: BLE001  # show any pipeline failure as a readable message
        st.error(f"미리보기에 실패했습니다: {e}")
        return
    _show_result(result, clean_rows(edited))
