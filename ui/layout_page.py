"""📄 발주서 양식: edit the columns of each vendor's purchase-order Excel file (one form at a time)."""
from __future__ import annotations

import pandas as pd
import requests
import streamlit as st

from store.base import ConflictError, DataStore, StoreError
from store.dictionary_repo import now_kst_iso
from store.layout_repo import OUTPUT_LAYOUT_FILE
from ui import context, layout_state
from ui.layout_logic import (
    BLANK_LABEL,
    NEW_COLUMN_LABEL,
    RID,
    V_DELETE,
    V_FIXED,
    V_NAME,
    V_ORDER,
    V_SOURCE,
    Judgement,
    build_view,
    check_new_form,
    column_choices,
    fixed_options,
    form_filename,
    form_names,
    judge,
    preview_table,
    set_one_column,
    source_options,
    summary_message,
    used_sources,
)
from ui.layout_state import LayoutData

HELP = (
    "'순서'는 엑셀 파일에서 왼쪽부터 몇 번째 칸인지를 뜻합니다. 저장하면 열(A, B, C…)이 순서대로 다시 매겨집니다. "
    "'고정 글자'를 적으면 모든 줄에 그 글자가 들어가고(넣을 내용보다 우선), 비워 두면 '넣을 내용'이 들어갑니다. "
    "표 안에서는 한글이 잘 입력되지 않으니, 한글 칸 이름이나 새 고정 글자는 표 아래 '칸 하나 추가·고치기'에서 넣으세요."
)
DIRECT_INPUT = "✏️ 직접 입력"
NO_FIXED = "(없음)"
S_FORM = "layout_w_form"


# ---------------------------------------------------------------- saving

def _write(store: DataStore, data: LayoutData, text: str, message: str, mine_key: str) -> None:
    try:
        store.write_text(OUTPUT_LAYOUT_FILE, text, data.sha, message, context.current_author())
    except ConflictError:
        st.session_state[layout_state.S_CONFLICT] = (mine_key, text)
        return
    except (StoreError, requests.RequestException) as e:
        st.error(f"저장하지 못했습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return
    context.clear_data_caches()
    layout_state.load(store)
    st.session_state[layout_state.S_FLASH] = f"저장했습니다. {message}"
    st.rerun()


def _conflict_box(form: str) -> None:
    saved = st.session_state.get(layout_state.S_CONFLICT)
    if saved is None or saved[0] != form:
        return
    st.error("다른 사람이 먼저 저장했습니다. '새로 불러오기' 후 다시 수정하세요.")
    st.download_button(
        "내 수정본 받기", data=saved[1].encode("utf-8"), file_name="output_layout_내수정본.csv",
        mime="text/csv", key="layout_w_mine",
    )


def _issue_lists(judged: Judgement) -> None:
    for line in judged.errors:
        st.error(line)
    if judged.warnings:
        with st.expander(f"⚠ 고친 칸의 경고 {len(judged.warnings)}건 (저장 전 확인)"):
            for line in judged.warnings:
                st.warning(line)
    if judged.reference:
        with st.expander(f"참고: 이번에 고치지 않은 부분의 문제 {len(judged.reference)}건 (저장은 막지 않음)"):
            for line in judged.reference:
                st.caption(line)


def _save_panel(store: DataStore, data: LayoutData, form: str, judged: Judgement) -> None:
    edit = judged.edit
    st.caption(f"수정 {edit.modified} · 추가 {edit.added} · 삭제 {edit.deleted}")
    clicked = st.button("저장", type="primary", disabled=not judged.has_changes, key="layout_w_save")
    _issue_lists(judged)
    ask = bool(judged.warnings) and not judged.errors
    confirmed = st.checkbox("경고를 확인했습니다", key="layout_w_confirm") if ask else False
    if not clicked:
        return
    if judged.errors:
        st.error("오류가 있어 저장하지 않았습니다.")
    elif judged.warnings and not confirmed:
        st.warning("경고를 확인한 뒤 '경고를 확인했습니다'에 체크하고 다시 저장하세요.")
    else:
        _write(store, data, edit.text, summary_message(form, edit), form)


def _delete_panel(store: DataStore, data: LayoutData, form: str, view0: pd.DataFrame, filename: str) -> None:
    """Remove the whole form (the validator blocks it while rules or dictionary rows still use it)."""
    st.markdown("**이 양식 삭제**")
    confirmed = st.checkbox("이 양식을 통째로 삭제합니다", key="layout_w_delete_confirm")
    if not st.button("이 양식 삭제", disabled=not confirmed, key="layout_w_delete"):
        return
    judged = judge(
        data.table, data.text, form, filename, view0, view0.iloc[0:0], data.refs, data.route_csv,
        data.options_csv, context.current_user(), now_kst_iso(),
    )
    if judged.errors:
        for line in judged.errors:
            st.error(line)
        return
    _write(store, data, judged.edit.text, f"발주서 양식 삭제: {form}", form)


# ---------------------------------------------------------------- grid

def _frozen(data: LayoutData, form: str) -> dict:
    """The rows the editor was opened with (view0) and the rows it currently starts from (base).

    view0 stays frozen while the page reruns, so edits made in the grid are not replayed onto changing
    data; a reload, a save or another form gives a fresh editor. The input box under the grid changes
    `base` and bumps `nonce`, which gives a fresh editor that starts from the changed rows.
    """
    version = st.session_state[layout_state.S_VERSION]
    frozen = st.session_state.get(layout_state.S_FROZEN)
    if frozen is None or frozen["sig"] != (version, form):
        view0 = build_view(data.table, form)
        frozen = {"sig": (version, form), "view0": view0, "base": view0, "nonce": 0}
        st.session_state[layout_state.S_FROZEN] = frozen
    return frozen


def _grid(data: LayoutData, form: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """(frozen rows the editor was opened with, the editor's current result)."""
    frozen = _frozen(data, form)
    version, nonce = st.session_state[layout_state.S_VERSION], frozen["nonce"]
    key = f"layout_w_grid_{version}_{form}" + (f"_{nonce}" if nonce else "")
    edited = st.data_editor(
        frozen["base"], key=key, num_rows="dynamic", hide_index=True,
        column_config=_column_config(data, frozen["base"]), width="stretch",
    )
    return frozen["view0"], edited


def _column_config(data: LayoutData, base: pd.DataFrame) -> dict:
    return {
        RID: None,
        V_ORDER: st.column_config.NumberColumn("순서", step=1, format="%d", help="왼쪽에서 몇 번째 칸인지"),
        V_NAME: st.column_config.TextColumn("칸 이름", help="엑셀 첫 줄에 적히는 이름 (한글은 표 아래 입력란에서)"),
        V_SOURCE: st.column_config.SelectboxColumn(
            "넣을 내용", options=source_options(used_sources(data.table)), default=BLANK_LABEL,
            help="주문 데이터에서 가져올 내용 (고정 글자가 있으면 무시됨)",
        ),
        V_FIXED: st.column_config.SelectboxColumn(
            "고정 글자", options=fixed_options(data.table, base), default="",
            help="고르면 모든 줄에 이 글자가 들어갑니다. 목록에 없는 글자는 표 아래 입력란에서 직접 입력",
        ),
        V_DELETE: st.column_config.CheckboxColumn("삭제", default=False),
    }


def _text(value: object) -> str:
    return "" if value is None or pd.isna(value) else str(value).strip()


def _one_column_box(data: LayoutData, form: str, edited: pd.DataFrame) -> None:
    """Add or change one column with normal input boxes (Korean works here, unlike in the grid)."""
    frozen = _frozen(data, form)
    with st.expander("칸 하나 추가·고치기 (한글 입력은 여기서)", expanded=False):
        choices = column_choices(edited)
        target = st.selectbox("어느 칸?", choices, key=f"layout_w_one_target_{form}")
        index = None if target == NEW_COLUMN_LABEL else choices.index(target) - 1
        row = None if index is None else edited.iloc[index]
        suffix = f"{frozen['nonce']}_{form}_{target}"
        name = st.text_input("칸 이름", value="" if row is None else _text(row[V_NAME]), key=f"layout_w_one_name_{suffix}")
        sources = source_options(used_sources(data.table))
        current_source = (_text(row[V_SOURCE]) if row is not None else "") or BLANK_LABEL
        source = st.selectbox(
            "넣을 내용", sources, index=sources.index(current_source) if current_source in sources else 0,
            key=f"layout_w_one_source_{suffix}",
        )
        fixed = _fixed_input(data, edited, "" if row is None else _text(row[V_FIXED]), suffix)
        if st.button("표에 반영", key=f"layout_w_one_apply_{form}"):
            if not name.strip():
                st.error("칸 이름을 입력하세요.")
                return
            frozen["base"] = set_one_column(edited, index, name, source, fixed)
            frozen["nonce"] += 1
            st.rerun()
        st.caption("'표에 반영'을 눌러도 아직 저장되지 않습니다. 표와 미리보기를 확인한 뒤 '저장'을 누르세요.")


def _fixed_input(data: LayoutData, edited: pd.DataFrame, current: str, suffix: str) -> str:
    """고정 글자: pick one already used, or type a new one."""
    presets = [v for v in fixed_options(data.table, edited) if v]
    choices = [NO_FIXED, *presets, DIRECT_INPUT]
    start = choices.index(current) if current in presets else (0 if not current else len(choices) - 1)
    picked = st.selectbox("고정 글자", choices, index=start, key=f"layout_w_one_fixed_{suffix}",
                          help="고르면 모든 줄에 이 글자가 들어갑니다 (넣을 내용보다 우선)")
    if picked == DIRECT_INPUT:
        return st.text_input("고정 글자 직접 입력", value=current, key=f"layout_w_one_fixed_text_{suffix}")
    return "" if picked == NO_FIXED else picked


def _preview(edited: pd.DataFrame) -> None:
    st.markdown("**미리보기** (엑셀 첫 줄 + 예시 한 줄)")
    shown = preview_table(edited)
    if shown.columns.empty:
        st.caption("칸이 없습니다.")
    else:
        st.dataframe(shown, hide_index=True, width="stretch")


# ---------------------------------------------------------------- form choice

def _pending() -> dict[str, str]:
    return st.session_state.setdefault(layout_state.S_PENDING, {})


def _add_form() -> None:
    data: LayoutData | None = st.session_state.get(layout_state.S_DATA)
    name = str(st.session_state.get("layout_w_new_name", ""))
    filename = str(st.session_state.get("layout_w_new_file", ""))
    existing = [*form_names(data.table), *_pending()] if data else list(_pending())
    problem = check_new_form(name, filename, existing)
    if problem:
        st.session_state["layout_w_new_error"] = problem
        return
    st.session_state.pop("layout_w_new_error", None)
    _pending()[name.strip()] = filename.strip() or name.strip()
    st.session_state[S_FORM] = name.strip()


def _new_form_box() -> None:
    with st.expander("새 양식 추가"):
        st.text_input("양식 이름", key="layout_w_new_name")
        st.text_input("파일 이름 (비우면 양식 이름)", key="layout_w_new_file")
        st.button("양식 추가", on_click=_add_form, key="layout_w_new_add")
        error = st.session_state.get("layout_w_new_error")
        if error:
            st.error(error)
        st.caption("추가한 양식은 칸을 1개 이상 만들고 저장해야 만들어집니다. 상품분류·품목 사전에서 쓰려면 저장한 뒤 연결하세요.")


def _edit_form(store: DataStore, data: LayoutData, form: str) -> None:
    is_new = form in _pending() and form not in form_names(data.table)
    default = _pending()[form] if is_new else form_filename(data.table, form)
    filename = st.text_input("파일 이름", value=default, key=f"layout_w_file_{st.session_state[layout_state.S_VERSION]}_{form}")
    st.caption(HELP)
    view0, edited = _grid(data, form)
    _one_column_box(data, form, edited)
    _preview(edited)
    judged = judge(
        data.table, data.text, form, filename, view0, edited, data.refs, data.route_csv, data.options_csv,
        context.current_user(), now_kst_iso(),
    )
    _save_panel(store, data, form, judged)
    _conflict_box(form)
    if is_new:
        if st.button("이 새 양식 취소", key="layout_w_cancel_new"):
            _pending().pop(form, None)
            st.rerun()
    else:
        _delete_panel(store, data, form, view0, filename)


# ---------------------------------------------------------------- page

def _read_only_fallback() -> None:
    st.info(
        "아직 양식이 데이터 저장소로 옮겨지지 않았습니다. 관리자가 이관 도구"
        "(python -m tools.migrate_layout)를 먼저 실행해야 이 화면에서 고칠 수 있습니다. "
        "지금은 config.xlsx의 양식을 읽기만 합니다."
    )
    layout = context.load_raw_layout()
    st.dataframe(layout, hide_index=True, width="stretch")


def _load(store: DataStore) -> bool:
    """Load into the session when needed; False (after an error message) when that failed."""
    if st.button("새로 불러오기", key="layout_w_reload") or layout_state.S_DATA not in st.session_state:
        try:
            layout_state.load(store)
        except (StoreError, requests.RequestException, ValueError) as e:
            st.error(f"양식을 불러올 수 없습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
            return False
    return True


def render() -> None:
    st.subheader("📄 발주서 양식")
    store = context.get_store()
    if store is None:
        st.info("데이터 저장소가 설정되지 않아 양식을 편집할 수 없습니다. 설정(secrets)을 확인하세요.")
        return
    if not _load(store):
        return
    flash = st.session_state.pop(layout_state.S_FLASH, None)
    if flash:
        st.success(flash)
    data: LayoutData | None = st.session_state[layout_state.S_DATA]
    if data is None:
        _read_only_fallback()
        return
    names = [*form_names(data.table), *(n for n in _pending() if n not in form_names(data.table))]
    if not names:
        st.info("양식이 없습니다. '새 양식 추가'로 만드세요.")
        _new_form_box()
        return
    if st.session_state.get(S_FORM) not in names:
        st.session_state[S_FORM] = names[0]
    form = st.selectbox("양식", names, key=S_FORM)
    _new_form_box()
    _edit_form(store, data, str(form))
