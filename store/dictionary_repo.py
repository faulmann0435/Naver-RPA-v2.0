"""Item dictionary and its settings, stored in the data repository."""
from __future__ import annotations

from core.dictionary import DictionarySettings, ItemDictionary
from store.base import DataStore
from store.csv_codec import BOM, from_csv_text
from store.rules_repo import DICTIONARY_FILE, SETTINGS_FILE


def load_settings(store: DataStore) -> DictionarySettings:
    """Read settings.json; a missing file gives the defaults."""
    snapshot = store.read_text(SETTINGS_FILE)
    if snapshot is None:
        return DictionarySettings()
    return DictionarySettings.from_json_text(snapshot.content)


def load_dictionary(store: DataStore) -> tuple[ItemDictionary, str | None]:
    """Read dictionary.csv as text. Returns (dictionary, sha); missing file -> (empty, None)."""
    snapshot = store.read_text(DICTIONARY_FILE)
    if snapshot is None:
        return ItemDictionary.empty(), None
    if not snapshot.content.strip(BOM + " \r\n"):
        return ItemDictionary.empty(), snapshot.sha
    frame = from_csv_text(snapshot.content, as_text=True)
    return ItemDictionary(frame, load_settings(store)), snapshot.sha
