import pandas as pd

from core.dictionary import DictionarySettings
from store.csv_codec import BOM, from_csv_text, to_csv_text
from store.dictionary_repo import load_dictionary, load_settings
from store.memory_store import MemoryStore
from store.rules_repo import DICTIONARY_COLUMNS, DICTIONARY_FILE, SETTINGS_FILE


def _csv(*rows: dict) -> str:
    frame = pd.DataFrame([{**dict.fromkeys(DICTIONARY_COLUMNS, ""), **r} for r in rows], columns=DICTIONARY_COLUMNS)
    return to_csv_text(frame)


def test_missing_files_give_defaults():
    store = MemoryStore()
    dictionary, sha = load_dictionary(store)
    assert len(dictionary) == 0 and sha is None
    assert load_settings(store) == DictionarySettings()


def test_round_trip_and_int_product_no_match():
    store = MemoryStore()
    text = _csv({
        "channel": "naver", "enabled": "1", "product_no": "5568579375", "option_key": "가리비 선택: 국산참가리비",
        "vendor_id": "V1", "display_template": "가리비 {수량}개", "unit_weight_kg": "0.5",
    })
    sha = store.write_text(DICTIONARY_FILE, text, None, "seed")
    dictionary, got_sha = load_dictionary(store)
    assert got_sha == sha and len(dictionary) == 1
    order_option = "수령일 선택 (도착시간 지정불가): 2월 5일 수령 / 가리비 선택: ⭐국산참가리비"
    entry = dictionary.lookup(5568579375, order_option, DictionarySettings())
    assert entry is not None and entry.unit_weight_kg == 0.5 and entry.vendor_id == "V1"


def test_header_only_file_is_empty_dictionary():
    store = MemoryStore()
    store.write_text(DICTIONARY_FILE, to_csv_text(pd.DataFrame(columns=DICTIONARY_COLUMNS)), None, "seed")
    dictionary, sha = load_dictionary(store)
    assert len(dictionary) == 0 and sha is not None


def test_settings_loaded_from_store():
    store = MemoryStore()
    store.write_text(SETTINGS_FILE, '{"item_separator": " | "}', None, "seed")
    assert load_settings(store).item_separator == " | "


def test_from_csv_text_as_text_keeps_strings():
    text = BOM + "a,b\n0012,\n"
    assert from_csv_text(text, as_text=True).iloc[0].tolist() == ["0012", ""]
    assert from_csv_text(text)["a"].iloc[0] == 12
