"""Streamlit-aware shared services: identity, data store, config, dictionary state, caches."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import requests
import streamlit as st

from core.config_loader import load_config, read_config_sheets
from core.dictionary import DEFAULT_WORKERS, DictionarySettings, ItemDictionary
from store.base import Author, DataStore, StoreError
from store.dictionary_repo import empty_frame, load_dictionary_frame, load_state
from store.github_store import GitHubStore
from store.layout_repo import load_layout
from store.rules_repo import load_config_from_store

CONFIG_PATH = str(Path(__file__).resolve().parent.parent / "config.xlsx")
CONFIG_PASSWORD = "1111"
WORKER_KEY = "worker_name"
WORKER_PARAM = "worker"
AUTHOR_EMAIL = "app@naver-rpa.local"
TEST_STORE_KEY = "_test_store"
CACHE_SECONDS = 60

LAYOUT_SOURCE_STORE, LAYOUT_SOURCE_FILE = "데이터 저장소", "config.xlsx"

DictionaryState = tuple[ItemDictionary, DictionarySettings, str | None]


# ---------------------------------------------------------------- identity

def worker_names() -> tuple[str, ...]:
    """Worker names from settings.json; the default list when there is no store or it fails."""
    try:
        return load_dictionary_state()[1].workers
    except (StoreError, requests.RequestException, KeyError, ValueError):
        return DEFAULT_WORKERS


def _store_worker_choice() -> None:
    st.query_params[WORKER_PARAM] = str(st.session_state[WORKER_KEY])


def _worker_picker(workers: tuple[str, ...]) -> None:
    if st.session_state.get(WORKER_KEY) not in workers:
        saved = st.query_params.get(WORKER_PARAM)
        st.session_state[WORKER_KEY] = saved if saved in workers else workers[0]
    st.sidebar.selectbox("작업자", workers, key=WORKER_KEY, on_change=_store_worker_choice)
    st.sidebar.caption("저장할 때 이 이름이 변경 이력에 남습니다")


def current_user(show_widget: bool = False) -> str:
    """The worker's name. The ONLY place that knows who is working.

    The app cannot know who is logged in, so the worker picks their name in the sidebar.
    """
    workers = worker_names()
    if show_widget:
        _worker_picker(workers)
    chosen = st.session_state.get(WORKER_KEY)
    return str(chosen) if chosen in workers else workers[0]


def current_author() -> Author:
    return Author(name=current_user(), email=AUTHOR_EMAIL)


# ---------------------------------------------------------------- data store

def _data_store_settings() -> tuple[str, str] | None:
    """(repo, branch) from secrets, or None when no data store is configured."""
    try:
        if "data_store" not in st.secrets:
            return None
        section = st.secrets["data_store"]
        return str(section["repo"]), str(section["branch"])
    except Exception:  # noqa: BLE001  # no secrets file / missing keys
        return None


def _github_store(repo: str, branch: str) -> GitHubStore:
    """The token is read here and never passed around (or used as a cache key)."""
    return GitHubStore(repo=repo, branch=branch, token=st.secrets["data_store"]["token"])


def get_store() -> DataStore | None:
    injected = st.session_state.get(TEST_STORE_KEY)
    if injected is not None:
        return injected  # type: ignore[no-any-return]
    settings = _data_store_settings()
    if settings is None:
        return None
    try:
        return _github_store(*settings)
    except Exception:  # noqa: BLE001  # token missing in the secrets section
        return None


def store_label() -> str:
    """Branch name of the data store (dev / main), shown in the sidebar."""
    if st.session_state.get(TEST_STORE_KEY) is not None:
        return "test"
    settings = _data_store_settings()
    return settings[1] if settings else ""


# ---------------------------------------------------------------- config

@st.cache_data(ttl=3600)
def load_config_local(config_path: str, _password: str | None = None, cache_key: str | None = None):
    """Cached wrapper around core.config_loader.load_config (cache_key: file mtime/size)."""
    return load_config(config_path, password=_password)


@st.cache_data(ttl=CACHE_SECONDS)
def _load_config_from_data_store(repo: str, branch: str, config_path: str, cache_key: str) -> dict:
    return load_config_from_store(_github_store(repo, branch), config_path, password=CONFIG_PASSWORD)


def _file_cache_key(config_path: str) -> str:
    stat = Path(config_path).stat()
    return f"{stat.st_mtime}_{stat.st_size}"


def get_config(config_path: str = CONFIG_PATH, cache_key: str | None = None, warn: bool = True) -> tuple[dict, str]:
    """Rules from the data repository when configured, else from config.xlsx (with a warning)."""
    key = cache_key if cache_key is not None else _file_cache_key(config_path)
    injected = st.session_state.get(TEST_STORE_KEY)
    if injected is not None:
        return load_config_from_store(injected, config_path, password=CONFIG_PASSWORD), "테스트 저장소"
    settings = _data_store_settings()
    if settings is not None:
        repo, branch = settings
        try:
            return _load_config_from_data_store(repo, branch, config_path, key), f"데이터 저장소 ({repo}@{branch})"
        # Connection failures, missing secrets keys and malformed rule files must not stop
        # order processing: warn and fall back to the rules in config.xlsx.
        except (StoreError, requests.RequestException, KeyError, ValueError) as e:
            if warn:
                st.warning(
                    "데이터 저장소의 규칙을 불러올 수 없어 config.xlsx의 규칙으로 처리합니다. "
                    f"최신 규칙이 아닐 수 있습니다. (원인: {type(e).__name__}: {str(e)[:200]})"
                )
    return load_config_local(config_path, _password=CONFIG_PASSWORD, cache_key=key), "config.xlsx"


def load_config_or_error() -> dict | None:
    """Config for a page; shows a Korean error and returns None when it cannot be loaded."""
    try:
        return get_config(warn=False)[0]
    except FileNotFoundError:
        st.error(f"설정 파일을 찾을 수 없습니다. 경로: {CONFIG_PATH}")
    except (ValueError, StoreError) as e:
        st.error(str(e))
    return None


@st.cache_data(ttl=3600)
def _raw_layout(config_path: str, _password: str | None, cache_key: str) -> pd.DataFrame:
    return read_config_sheets(config_path, _password)["OutputLayout"]


@st.cache_data(ttl=CACHE_SECONDS)
def _cached_store_layout(repo: str, branch: str) -> pd.DataFrame | None:
    snapshot = load_layout(_github_store(repo, branch))
    return None if snapshot is None else snapshot.raw


def _store_layout() -> pd.DataFrame | None:
    """OutputLayout of the data store (None: no store, no output_layout.csv yet, or the store failed)."""
    injected = st.session_state.get(TEST_STORE_KEY)
    if injected is not None:
        snapshot = load_layout(injected)
        return None if snapshot is None else snapshot.raw
    settings = _data_store_settings()
    if settings is None:
        return None
    try:
        return _cached_store_layout(*settings)
    except (StoreError, requests.RequestException, KeyError, ValueError):
        return None  # same fallback as get_config: keep working with config.xlsx


def load_layout_with_source() -> tuple[pd.DataFrame, str]:
    """(OutputLayout as raw sheet frame, where it came from): the data store once migrated, else config.xlsx."""
    stored = _store_layout()
    if stored is not None:
        return stored.copy(), LAYOUT_SOURCE_STORE
    return _raw_layout(CONFIG_PATH, CONFIG_PASSWORD, _file_cache_key(CONFIG_PATH)).copy(), LAYOUT_SOURCE_FILE


def load_raw_layout() -> pd.DataFrame:
    """OutputLayout in raw sheet shape (the rule checks normalize it themselves)."""
    return load_layout_with_source()[0]


def vendor_ids(config: dict) -> list[str]:
    """Distinct VendorID values of OutputLayout, in sheet order."""
    values = (str(v).strip() for v in config["OutputLayout"]["VendorID"].dropna())
    return list(dict.fromkeys(v for v in values if v))


# ---------------------------------------------------------------- dictionary

@st.cache_data(ttl=CACHE_SECONDS)
def _cached_dictionary_state(repo: str, branch: str) -> DictionaryState:
    return load_state(_github_store(repo, branch))


def load_dictionary_state() -> DictionaryState:
    """(dictionary, settings, sha); cached for 60 s with the real store."""
    injected = st.session_state.get(TEST_STORE_KEY)
    if injected is not None:
        return load_state(injected)
    settings = _data_store_settings()
    if settings is None:
        return ItemDictionary.empty(), DictionarySettings(), None
    return _cached_dictionary_state(*settings)


def load_dictionary_state_or_error() -> DictionaryState | None:
    try:
        return load_dictionary_state()
    except (StoreError, requests.RequestException, KeyError, ValueError) as e:
        st.error(f"품목 사전을 불러올 수 없습니다. (원인: {type(e).__name__}: {str(e)[:200]})")
        return None


def load_dictionary_frame_fresh() -> tuple[pd.DataFrame, str | None]:
    """Raw dictionary.csv rows (text) and sha, never cached."""
    store = get_store()
    if store is None:
        return empty_frame(), None
    return load_dictionary_frame(store)


def clear_data_caches() -> None:
    """Call after any save so the next read sees the new data."""
    _cached_dictionary_state.clear()
    _load_config_from_data_store.clear()
    _cached_store_layout.clear()


# ---------------------------------------------------------------- sidebar

def _layout_source_label() -> str:
    try:
        return load_layout_with_source()[1]
    except (OSError, ValueError):
        return LAYOUT_SOURCE_FILE


def render_sidebar(config_source: str) -> None:
    current_user(show_widget=True)
    with st.sidebar:
        st.subheader("Config")
        st.caption(f"규칙 출처: {config_source}")
        st.caption(f"양식 출처: {_layout_source_label()}")
        if get_store() is not None:
            try:
                dictionary, _settings, _sha = load_dictionary_state()
                st.caption(f"사전: {len(dictionary)}개 항목 ({store_label()})")
            except (StoreError, requests.RequestException, KeyError, ValueError):
                st.caption("사전을 불러올 수 없습니다")
