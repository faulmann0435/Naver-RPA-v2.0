"""Session keys and loading of the 고급 설정 page (streamlit-aware, shared by its tabs)."""
from __future__ import annotations

import streamlit as st

from store.base import DataStore
from store.rules_repo import SETTINGS_FILE
from ui.rules_logic import OPTIONS_SPEC, ROUTE_SPEC, TableSpec, TableState, load_table

S_ROUTE, S_OPTIONS, S_SETTINGS = "rules_data_route", "rules_data_options", "rules_data_settings"
S_VERSION, S_FLASH = "rules_ver", "rules_flash"
FROZEN_PREFIX, CONFLICT_PREFIX = "rules_frozen_", "rules_conflict_"
SESSION_PREFIXES = ("rules_data_", FROZEN_PREFIX, CONFLICT_PREFIX)  # loaded data; widgets use "rules_w_"


def state_key(spec: TableSpec) -> str:
    return S_ROUTE if spec.kind == ROUTE_SPEC.kind else S_OPTIONS


def load_all(store: DataStore) -> None:
    """(Re)load both rule files and settings.json into the session; raises StoreError."""
    tables = {S_ROUTE: load_table(store, ROUTE_SPEC), S_OPTIONS: load_table(store, OPTIONS_SPEC)}
    snapshot = store.read_text(SETTINGS_FILE)
    version = st.session_state.get(S_VERSION, 0) + 1  # new editor keys: old edit state is never replayed
    forget_loaded()
    st.session_state.update(tables)
    st.session_state[S_SETTINGS] = (snapshot.content, snapshot.sha) if snapshot else (None, None)
    st.session_state[S_VERSION] = version


def forget_loaded() -> None:
    """Make the page reload from the store on its next run (used after a revert)."""
    for key in [k for k in st.session_state if str(k).startswith(SESSION_PREFIXES)]:
        del st.session_state[key]


def tables() -> tuple[TableState, TableState]:
    return st.session_state[S_ROUTE], st.session_state[S_OPTIONS]

