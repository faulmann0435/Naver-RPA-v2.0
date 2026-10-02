"""📦 주문처리: upload an order file, process it, register unmatched items, download purchase orders."""
from __future__ import annotations

import pandas as pd
import requests
import streamlit as st

from core.dictionary import DictionarySettings, ItemDictionary
from core.loader import HAS_MSOFFCRYPTO, load_excel, read_csv_with_encoding
from core.merger import filter_instruction_rows
from core.pipeline import ProcessResult, process_orders
from store.base import ConflictError, StoreError
from ui import context
from ui.option_form import (
    clear_pending_with_prefix,
    option_form,
    pending_values,
    state_key,
)
from ui.option_logic import FormValues, form_issues
from ui.order_logic import (
    RegisterResult,
    UnmatchedItem,
    build_suggestions,
    item_prefix,
    register_items,
    suggestion_values,
    unmatched_items,
)

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
REQUIRED_COLUMNS = ["수량", "상품명", "옵션정보", "배송비 묶음번호"]
S_CHECKED, K_MORE, PAGE_SIZE = "um_checked", "um_more", 10
S_RESULT, S_DF, S_SUGGEST, S_FILE, S_FLASH = (
    "order_result", "order_df", "order_suggestions", "last_file_id", "order_flash",
)


# ---------------------------------------------------------------- loading

def _read_upload(uploaded_file) -> pd.DataFrame | None:
    password = None
    name = (uploaded_file.name or "").lower()
    if name.endswith(".xlsx"):
        password = st.text_input("주문 엑셀 비밀번호 (없으면 비움)", type="password", key="order_pw")
    try:
        if name.endswith(".xlsx"):
            return load_excel(uploaded_file, password=password)
        uploaded_file.seek(0)
        return read_csv_with_encoding(uploaded_file)
    except Exception as e:  # noqa: BLE001  # any parsing/decryption failure becomes a readable message
        st.error(str(e))
        return None


def _prepare(df: pd.DataFrame) -> pd.DataFrame | None:
    before = len(df)
    df = filter_instruction_rows(df)
    if before > len(df):
        st.info(f"안내 문구 행 {before - len(df)}개 제거.")
    if df.empty:
        st.warning("처리할 데이터가 없습니다.")
        return None
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        st.error(f"필수 컬럼 누락: {', '.join(str(x) for x in missing)}")
        return None
    return df


# ---------------------------------------------------------------- processing

def _reset_unmatched_state() -> None:
    clear_pending_with_prefix("um_")
    st.session_state.pop(S_CHECKED, None)
    st.session_state.pop(K_MORE, None)


def _dictionary_args() -> tuple[ItemDictionary, DictionarySettings] | None:
    state = context.load_dictionary_state_or_error()
    return None if state is None else (state[0], state[1])


def _process_and_store(df: pd.DataFrame, config: dict) -> bool:
    args = _dictionary_args()
    if args is None:
        return False
    dictionary, settings = args
    try:
        result = process_orders(df, config, dictionary, settings)
    except Exception as e:  # noqa: BLE001  # surface any processing failure to the user
        st.error(f"처리 실패: {e}")
        return False
    suggestions = build_suggestions(result.unmatched, config, context.vendor_ids(config), settings)
    st.session_state[S_RESULT] = result
    st.session_state[S_DF] = df
    st.session_state[S_SUGGEST] = suggestions
    _reset_unmatched_state()
    return True


# ---------------------------------------------------------------- result sections

def _summary(result: ProcessResult) -> None:
    stats = result.stats
    cols = st.columns(4)
    cols[0].metric("사전 적용", stats["rows_dictionary"])
    cols[1].metric("자동 규칙", stats["rows_auto"])
    cols[2].metric("미분류", stats["rows_unclassified"])
    cols[3].metric("검토 필요 사전 항목", stats["rows_needs_review"])


def _show_register_outcome(result: RegisterResult) -> None:
    if result.status == "nothing_selected":
        st.info("등록할 항목을 먼저 '등록'에 체크하세요.")
    if result.errors:
        st.error("저장하지 않았습니다. 아래 오류를 고치세요.\n\n" + "\n".join(f"- {m}" for m in result.errors))
    if result.status == "needs_confirm":
        st.warning("경고가 있습니다. 확인 후 '경고를 확인했습니다'에 체크하고 다시 누르세요.")


def _register_and_reprocess(
    selected: list[tuple[UnmatchedItem, FormValues]], config: dict, ids: list[str], allow_warnings: bool
) -> None:
    store = context.get_store()
    args = _dictionary_args()
    if store is None or args is None:
        return
    try:
        outcome = register_items(store, selected, context.current_author(), args[1], ids, allow_warnings)
    except ConflictError:
        st.error("다른 사람이 동시에 사전을 수정해 저장하지 못했습니다. 잠시 후 다시 시도하세요.")
        return
    except (StoreError, requests.RequestException) as e:
        st.error(f"사전을 저장하지 못했습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    _show_register_outcome(outcome)
    if outcome.status != "saved" or outcome.appended is None:
        return
    context.clear_data_caches()
    if _process_and_store(st.session_state[S_DF], config):
        added, skipped = len(outcome.appended.added), len(outcome.appended.skipped_existing)
        st.session_state[S_FLASH] = (
            f"사전에 {added}건 등록했습니다. 이미 등록되어 건너뛴 항목: {skipped}건. 다시 처리했습니다."
        )
        st.rerun()


def _register_controls(
    selected: list[tuple[UnmatchedItem, FormValues]], config: dict, ids: list[str]
) -> None:
    settings = (_dictionary_args() or (None, DictionarySettings()))[1]
    warnings = [
        issue.message
        for item, values in selected
        for issue in form_issues(values, ids, item.qty_example, settings)
        if issue.level == "warning"
    ]
    for message in dict.fromkeys(warnings):
        st.warning(message)
    confirmed = st.checkbox("경고를 확인했습니다", key="order_confirm_warnings") if warnings else False
    if st.button("선택한 항목 사전에 등록하고 다시 처리"):
        _register_and_reprocess(selected, config, ids, confirmed)


def _unmatched_item(item: UnmatchedItem, initial: FormValues, ids: list[str]) -> tuple[bool, FormValues]:
    prefix = item_prefix(item)
    check_key = state_key(prefix, "register")
    st.session_state.setdefault(check_key, prefix in st.session_state.get(S_CHECKED, frozenset()))
    with st.container(border=True):
        st.markdown(
            f"**{item.name}** | {item.option or '옵션 없음'} | 수량 예시 {item.qty_example} | {item.count}건"
        )
        checked = st.checkbox("등록", key=check_key)
        values = option_form(prefix, initial, ids, show_vendor=True, qty_example=item.qty_example)
    return checked, values


def _unmatched_forms(
    suggestions: pd.DataFrame, ids: list[str]
) -> list[tuple[UnmatchedItem, FormValues]]:
    """One form per unmatched item (first PAGE_SIZE, the rest behind a toggle); the checked ones."""
    items = unmatched_items(suggestions)
    initials = [suggestion_values(suggestions, i) for i in range(len(items))]
    rendered: dict[str, bool] = {}
    for position, (item, initial) in enumerate(zip(items, initials, strict=True)):
        if position == PAGE_SIZE and not st.toggle(f"더 보기 (나머지 {len(items) - PAGE_SIZE}건)", key=K_MORE):
            break
        rendered[item_prefix(item)] = _unmatched_item(item, initial, ids)[0]
    old = st.session_state.get(S_CHECKED, frozenset())
    checked = frozenset(p for p, flag in rendered.items() if flag) | (old - set(rendered))
    st.session_state[S_CHECKED] = checked
    return [
        (item, pending_values(item_prefix(item)) or initial)
        for item, initial in zip(items, initials, strict=True) if item_prefix(item) in checked
    ]


def _unmatched_box(result: ProcessResult, config: dict) -> None:
    unmatched_count = len(result.unmatched)
    unclassified = result.stats["rows_unclassified"]
    if unmatched_count == 0 and unclassified == 0:
        return
    if unmatched_count:
        st.warning(f"⚠ 사전 미등록 {unmatched_count}건 — 자동 규칙으로 처리했습니다. 확인 후 등록하세요.")
    if unclassified:
        st.warning(f"미분류 {unclassified}건은 어느 발주서에도 들어가지 않았습니다. 상품 분류 규칙을 확인하세요.")
    if not unmatched_count:
        return
    ids = context.vendor_ids(config)
    st.caption("숫자가 수량에 따라 바뀌면 그 숫자에 체크하세요. 결과는 바로 아래에 보입니다.")
    selected = _unmatched_forms(st.session_state[S_SUGGEST], ids)
    if context.get_store() is None:
        st.caption("데이터 저장소가 설정되지 않아 등록할 수 없습니다")
        return
    _register_controls(selected, config, ids)


def _path_expander(result: ProcessResult) -> None:
    with st.expander("행별 처리 경로"):
        if result.unmatched.empty:
            st.caption("모든 행이 사전에서 처리되었습니다.")
            return
        table = result.unmatched[["상품명", "옵션정보", "count"]].rename(columns={"count": "건수"})
        st.dataframe(table.assign(처리="자동 규칙"), hide_index=True, width="stretch")


def _downloads(result: ProcessResult) -> None:
    st.success(f"처리 완료. 총 {len(result.files)}개 업체 파일을 다운로드하세요.")
    for i, pf in enumerate(result.files):
        st.download_button(
            label=f"📥 Download [{pf['vendor']}] File",
            data=pf["data"].getvalue(),
            file_name=pf["filename"],
            mime=EXCEL_MIME,
            key=f"dl_{pf['vendor']}_{i}",
        )


def _show_result(result: ProcessResult, config: dict) -> None:
    flash = st.session_state.pop(S_FLASH, None)
    if flash:
        st.success(flash)
    _summary(result)
    _unmatched_box(result, config)
    _downloads(result)
    _path_expander(result)


# ---------------------------------------------------------------- page

def render() -> None:
    st.subheader("📦 주문처리")
    st.session_state.setdefault(S_RESULT, None)
    st.session_state.setdefault(S_FILE, None)
    config = context.load_config_or_error()
    if config is None:
        return
    if not HAS_MSOFFCRYPTO:
        st.warning("비밀번호 보호 엑셀: `pip install msoffcrypto-tool`")
    st.write("주문 파일 업로드 후 Process 버튼을 클릭하세요.")
    uploaded_file = st.file_uploader("주문 파일 (.xlsx 또는 .csv)", type=["xlsx", "csv"], key="uploaded_file")

    if uploaded_file is not None:
        file_id = (uploaded_file.name, uploaded_file.size)
        if st.session_state[S_FILE] != file_id:
            for key in (S_DF, S_SUGGEST):
                st.session_state.pop(key, None)
            _reset_unmatched_state()
            st.session_state[S_RESULT] = None
            st.session_state[S_FILE] = file_id

    result: ProcessResult | None = st.session_state[S_RESULT]
    if uploaded_file is None:
        if result is None:
            st.info("주문 파일을 업로드한 뒤 'Process Orders' 버튼을 클릭하세요.")
        else:
            _show_result(result, config)
        return

    raw = _read_upload(uploaded_file)
    df = None if raw is None else _prepare(raw)
    if df is None:
        return
    st.subheader("Raw Data Preview (상위 5행)")
    st.dataframe(df.head())
    if st.button("Process Orders"):
        with st.spinner("Processing..."):
            done = _process_and_store(df, config)
        if done:
            st.rerun()
        return
    if result is None:
        st.caption("위 'Process Orders' 버튼을 클릭하면 처리 후 다운로드 버튼이 표시됩니다.")
    else:
        _show_result(result, config)
