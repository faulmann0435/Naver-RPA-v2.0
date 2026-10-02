"""
Sokcho Order Processing System v14.4
- Quantity Display Lock: prevents double quantity stamping (e.g., "x2 x2")
- GROUP_MULTIPLY: append (x{Qty}) only, no param word repeat
- Config: load_config_local, List-based, Session State
UI only; processing logic lives in the core/ package.
"""
from pathlib import Path

import requests
import streamlit as st

from core.config_loader import load_config
from core.loader import HAS_MSOFFCRYPTO, load_excel, read_csv_with_encoding
from core.merger import filter_instruction_rows
from core.pipeline import process_all_data
from store.base import StoreError
from store.github_store import GitHubStore
from store.rules_repo import load_config_from_store

# ============== Config Loader (v14.3: load_config_local + cache) ==============

@st.cache_data(ttl=3600)
def load_config_local(config_path: str, _password: str | None = None, cache_key: str | None = None):
    """
    Cached wrapper around core.config_loader.load_config.
    cache_key: pass file mtime/size to invalidate cache when file changes.
    """
    return load_config(config_path, password=_password)


@st.cache_data(ttl=60)
def _load_config_from_data_store(repo: str, branch: str, config_path: str, cache_key: str) -> dict:
    """Rules from the data repository; the token is read here and never passed as an argument."""
    token = st.secrets["data_store"]["token"]
    store = GitHubStore(repo=repo, branch=branch, token=token)
    return load_config_from_store(store, config_path, password="1111")


def _data_store_settings() -> tuple[str, str] | None:
    """(repo, branch) from secrets, or None when no data store is configured."""
    try:
        if "data_store" not in st.secrets:
            return None
        section = st.secrets["data_store"]
        return str(section["repo"]), str(section["branch"])
    except Exception:  # noqa: BLE001  # no secrets file / missing keys
        return None


def get_config(config_path: str, cache_key: str) -> tuple[dict, str]:
    """Load config from the data repository when configured, else from config.xlsx."""
    settings = _data_store_settings()
    if settings is not None:
        repo, branch = settings
        try:
            config = _load_config_from_data_store(repo, branch, config_path, cache_key)
            return config, f"데이터 저장소 ({repo}@{branch})"
        # Connection failures, missing secrets keys and malformed rule files must not stop
        # order processing: warn and fall back to the rules in config.xlsx.
        except (StoreError, requests.RequestException, KeyError, ValueError) as e:
            st.warning(
                "데이터 저장소의 규칙을 불러올 수 없어 config.xlsx의 규칙으로 처리합니다. "
                f"최신 규칙이 아닐 수 있습니다. (원인: {type(e).__name__}: {str(e)[:200]})"
            )
    return load_config_local(config_path, _password="1111", cache_key=cache_key), "config.xlsx"


# ============== UI (v14.3: Session State) ==============

def main():
    st.set_page_config(page_title="속초 발주 처리 시스템 v14.4", layout="wide")
    st.title("속초 발주 처리 시스템 v14.4")

    # Session state: persist processed results so download buttons do NOT trigger re-run
    if "processed_results" not in st.session_state:
        st.session_state.processed_results = None
    if "last_file_id" not in st.session_state:
        st.session_state.last_file_id = None

    st.write("설정: config.xlsx (로컬). 주문 파일 업로드 후 Process 버튼을 클릭하세요.")

    if not HAS_MSOFFCRYPTO:
        st.warning("비밀번호 보호 엑셀: `pip install msoffcrypto-tool`")

    config_path = "config.xlsx"
    try:
        path = Path(config_path).resolve()
        if not path.exists():
            st.error(f"설정 파일을 찾을 수 없습니다. 경로: {path!s}")
            st.info("config.xlsx를 앱과 같은 폴더에 두거나 경로를 확인하세요.")
            return
        cache_key = f"{path.stat().st_mtime}_{path.stat().st_size}"
        config, config_source = get_config(config_path, cache_key)
    except FileNotFoundError as e:
        st.error(str(e))
        return
    except ValueError as e:
        st.error(str(e))
        return

    with st.sidebar:
        st.subheader("Config")
        st.caption(f"규칙 출처: {config_source}")

    uploaded_file = st.file_uploader("주문 파일 (.xlsx 또는 .csv)", type=["xlsx", "csv"], key="uploaded_file")

    # When a new order file is uploaded (different name/size), reset processed_results
    if uploaded_file is not None:
        file_id = (uploaded_file.name, uploaded_file.size)
        if st.session_state.last_file_id != file_id:
            st.session_state.processed_results = None
            st.session_state.last_file_id = file_id

    # No file and no cached results -> ask to upload
    if uploaded_file is None and st.session_state.processed_results is None:
        st.info("주문 파일을 업로드한 뒤 'Process Orders' 버튼을 클릭하세요.")
        return

    # No file but we have results (e.g. after download click rerun) -> show download only
    if uploaded_file is None and st.session_state.processed_results is not None:
        st.success(f"처리 완료. 총 {len(st.session_state.processed_results)}개 업체 파일을 다운로드하세요.")
        for i, pf in enumerate(st.session_state.processed_results):
            st.download_button(
                label=f"📥 Download [{pf['vendor']}] File",
                data=pf["data"].getvalue(),
                file_name=pf["filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{pf['vendor']}_{i}",
            )
        return

    # File uploaded: load and validate
    password = None
    if (uploaded_file.name or "").lower().endswith(".xlsx"):
        password = st.text_input("주문 엑셀 비밀번호 (없으면 비움)", type="password", key="order_pw")
    file_name = (uploaded_file.name or "").lower()
    try:
        if file_name.endswith(".xlsx"):
            df = load_excel(uploaded_file, password=password)
        else:
            uploaded_file.seek(0)
            df = read_csv_with_encoding(uploaded_file)
    except Exception as e:
        st.error(str(e))
        return

    before = len(df)
    df = filter_instruction_rows(df)
    if before > len(df):
        st.info(f"안내 문구 행 {before - len(df)}개 제거.")
    if df.empty:
        st.warning("처리할 데이터가 없습니다.")
        return

    required = ["수량", "상품명", "옵션정보", "배송비 묶음번호"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"필수 컬럼 누락: {', '.join(str(x) for x in missing)}")
        return

    st.subheader("Raw Data Preview (상위 5행)")
    st.dataframe(df.head())

    # Process button: run pipeline and store in session state
    if st.button("Process Orders"):
        with st.spinner("Processing..."):
            try:
                result = process_all_data(df, config)
                st.session_state.processed_results = result
                st.session_state.last_file_id = (uploaded_file.name, uploaded_file.size)
            except Exception as e:
                st.error(f"처리 실패: {e}")
                return
        st.rerun()

    # If we already have results for this session, show download section
    if st.session_state.processed_results is not None:
        st.success(f"처리 완료. 총 {len(st.session_state.processed_results)}개 업체 파일을 다운로드하세요.")
        for i, pf in enumerate(st.session_state.processed_results):
            st.download_button(
                label=f"📥 Download [{pf['vendor']}] File",
                data=pf["data"].getvalue(),
                file_name=pf["filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{pf['vendor']}_{i}",
            )
    else:
        st.caption("위 'Process Orders' 버튼을 클릭하면 처리 후 다운로드 버튼이 표시됩니다.")


if __name__ == "__main__":
    main()
