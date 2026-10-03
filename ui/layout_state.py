"""Session keys and loading of the 발주서 양식 page (streamlit-aware)."""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd
import streamlit as st

from store.base import DataStore
from store.layout_repo import OUTPUT_LAYOUT_FILE
from store.layout_validators import LayoutReferences, collect_references
from store.rules_repo import DICTIONARY_FILE, OPTION_RULES_FILE, PRODUCT_ROUTE_FILE
from ui.layout_logic import read_layout_table

S_DATA, S_VERSION, S_FLASH = "layout_data", "layout_ver", "layout_flash"
S_PENDING, S_FROZEN, S_CONFLICT = "layout_pending", "layout_frozen", "layout_conflict"
SESSION_KEYS = (S_DATA, S_PENDING, S_FROZEN, S_CONFLICT)  # loaded data; widgets use the prefix "layout_w_"


@dataclass(frozen=True)
class LayoutData:
    """output_layout.csv as loaded, plus what the checks need from the rule files (read once)."""

    text: str
    sha: str
    table: pd.DataFrame
    route_csv: str
    options_csv: str
    refs: LayoutReferences


def _text_or_empty(store: DataStore, path: str) -> str:
    snapshot = store.read_text(path)
    return "" if snapshot is None else snapshot.content


def load(store: DataStore) -> LayoutData | None:
    """(Re)load the layout and the rule files into the session; None when output_layout.csv does not exist yet."""
    snapshot = store.read_text(OUTPUT_LAYOUT_FILE)
    version = st.session_state.get(S_VERSION, 0) + 1  # new editor keys: old edit state is never replayed
    forget_loaded()
    st.session_state[S_VERSION] = version
    if snapshot is None:
        st.session_state[S_DATA] = None
        return None
    route, options = _text_or_empty(store, PRODUCT_ROUTE_FILE), _text_or_empty(store, OPTION_RULES_FILE)
    refs = collect_references(route, options, _text_or_empty(store, DICTIONARY_FILE))
    data = LayoutData(snapshot.content, snapshot.sha, read_layout_table(snapshot.content), route, options, refs)
    st.session_state[S_DATA] = data
    return data


def forget_loaded() -> None:
    """Make the page reload from the store on its next run (used after a revert)."""
    for key in SESSION_KEYS:
        st.session_state.pop(key, None)
