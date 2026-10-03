"""등록 시각 (created_at): set once at creation, kept on edits, optional in Excel, shown in both tabs."""
import dataclasses
import io
from datetime import datetime, timezone

import pandas as pd
import pytest

from core.dictionary import DictionarySettings
from store.base import Author
from store.csv_codec import from_csv_text, to_csv_text
from store.dictionary_repo import (
    add_missing_entries,
    append_entries,
    load_dictionary_frame,
)
from store.memory_store import MemoryStore
from store.rules_repo import DICTIONARY_COLUMNS, DICTIONARY_FILE
from tools.backfill_created_at import backfilled_frame
from tools.clean_dictionary_text import cleaned_frame
from tools.seed_build import SeedItem, dictionary_frame
from ui.dictionary_logic import (
    ADVANCED_VIEW,
    BASIC_VIEW,
    DISABLED_VIEW,
    EXPORT_HEADERS,
    V_CREATED,
    build_view,
    fold_edits,
    format_created,
    make_work,
    prepare_save,
    to_excel_bytes,
)
from ui.excel_import_logic import plan_import, read_import_frame
from ui.history_logic import diff_dictionary
from ui.option_logic import values_from_row
from ui.product_logic import (
    COL_CREATED,
    apply_product_edits,
    group_id,
    product_option_rows,
    product_summaries,
    summary_table,
)

SETTINGS = DictionarySettings()
AUTHOR = Author("tester", "t@example.com")
OLD = "2026-01-05T08:00:00+09:00"
NOW = "2026-10-03T09:54:24+09:00"


def _row(no: str, key: str = "a", name: str = "사과", **kw: str) -> dict[str, str]:
    base = dict.fromkeys(DICTIONARY_COLUMNS, "")
    return {**base, "channel": "naver", "enabled": "1", "needs_review": "0", "append_to_end": "0", "source": "manual",
            "product_no": no, "option_key": key, "product_name_ref": name, "option_raw_ref": key, "vendor_id": "V1",
            "display_template": f"{name} {{수량}}개", "updated_at": OLD, "updated_by": "old", **kw}


def _frame(*rows: dict[str, str]) -> pd.DataFrame:
    return pd.DataFrame(list(rows), columns=DICTIONARY_COLUMNS)


def test_created_at_is_the_last_column():
    assert DICTIONARY_COLUMNS[-1] == "created_at" and EXPORT_HEADERS["created_at"] == "등록 시각"


def test_format_created():
    assert format_created("2026-10-03T09:54:24+09:00") == "2026-10-03 09:54"
    assert format_created("") == "" and format_created(None) == "" and format_created("garbage") == ""


def test_register_from_order_page_stamps_created_at():
    store = MemoryStore()
    store.write_text(DICTIONARY_FILE, to_csv_text(_frame(_row("1"))), None, "init")
    append_entries(store, [{"product_no": "9", "option_key": "z", "vendor_id": "V1", "display_template": "Z"}],
                   AUTHOR, SETTINGS)
    frame, _ = load_dictionary_frame(store)
    assert frame.iloc[0]["created_at"] == "" and frame.iloc[1]["created_at"] == frame.iloc[1]["updated_at"] != ""


def test_old_csv_without_the_column_loads_as_empty():
    store = MemoryStore()
    old = _frame(_row("1")).drop(columns=["created_at"])
    store.write_text(DICTIONARY_FILE, to_csv_text(old), None, "init")
    frame, _ = load_dictionary_frame(store)
    assert list(frame.columns) == DICTIONARY_COLUMNS and frame.iloc[0]["created_at"] == ""


def test_seed_rows_and_merge_stamp_created_at():
    class _Entry:  # minimal stand-in for DictionaryEntry
        product_no, option_key, product_name_ref, option_raw_ref = "5", "k", "배", "k"
        vendor_id, display_template, sum_group, unit_weight_kg = "V1", "배 {수량}개", "", None

    seed = dictionary_frame([SeedItem(_Entry(), "auto_rule")], datetime(2026, 10, 3, 9, 0, tzinfo=timezone.utc))
    assert seed.iloc[0]["created_at"] == seed.iloc[0]["updated_at"] != ""
    store = MemoryStore()
    store.write_text(DICTIONARY_FILE, to_csv_text(_frame(_row("1"))), None, "init")
    add_missing_entries(store, seed.assign(created_at=""), AUTHOR, SETTINGS)
    frame, _ = load_dictionary_frame(store)
    assert frame.iloc[0]["created_at"] == "" and frame.iloc[1]["created_at"].endswith("+09:00")
    assert "created_at" in from_csv_text(to_csv_text(frame), as_text=True).columns


def test_import_new_row_is_stamped_and_updates_keep_created_at():
    current = _frame(_row("1", created_at=OLD))
    incoming = _frame(_row("1", display_template="새 표기", created_at="1999-01-01T00:00:00+09:00"),
                      _row("2", name="배", created_at="1999-01-01T00:00:00+09:00"))
    plan = plan_import(current, incoming, SETTINGS, "u", NOW)
    by_no = {r["product_no"]: r for _, r in plan.frame.iterrows()}
    assert by_no["1"]["created_at"] == OLD and by_no["1"]["updated_at"] == NOW
    assert by_no["2"]["created_at"] == NOW


def test_import_without_created_column_still_works_and_export_has_it():
    frame = _frame(_row("1", created_at=OLD))
    data = to_excel_bytes(frame)
    assert "등록 시각" in pd.read_excel(io.BytesIO(data)).columns
    buffer = io.BytesIO()
    pd.read_excel(io.BytesIO(data), dtype=str, keep_default_na=False).drop(columns=["등록 시각"]).to_excel(
        buffer, index=False)
    read = read_import_frame(buffer.getvalue())
    assert list(read.columns) == DICTIONARY_COLUMNS and read.iloc[0]["created_at"] == ""
    assert plan_import(frame, read, SETTINGS, "u", NOW).modified == 0


def test_edits_preserve_created_at():
    base = _frame(_row("1", created_at=OLD))
    work = make_work(base)
    work.loc[0, "display_template"] = "바뀜"
    assert prepare_save(base, work, "me", now=NOW).iloc[0]["created_at"] == OLD
    group = group_id("1", "사과")
    edit = dataclasses.replace(values_from_row(product_option_rows(base, group)[0]), template="바뀜")
    assert apply_product_edits(base, group, {("1", "a"): edit}, user="me", now=NOW).iloc[0]["created_at"] == OLD
    cleaned = cleaned_frame(_frame(_row("1", display_template="🔥x", created_at=OLD)), "c", NOW)[0]
    assert cleaned.iloc[0]["created_at"] == OLD


def test_grid_view_column_is_read_only_and_ignored_by_fold():
    assert V_CREATED == "등록일" and V_CREATED in DISABLED_VIEW
    work = make_work(_frame(_row("1", created_at=OLD), _row("2", name="배")))
    for advanced, columns in ((False, BASIC_VIEW), (True, ADVANCED_VIEW)):
        view = build_view(work, [0, 1], advanced)
        assert list(view.columns)[1:] == columns
        assert columns[columns.index("옵션") + 1] == V_CREATED
        assert list(view[V_CREATED]) == ["2026-01-05 08:00", ""]
        view.loc[0, V_CREATED] = "바꿔 봄"
        assert fold_edits(work, view).equals(work)


def test_summary_uses_earliest_created_at():
    frame = _frame(
        _row("1", "a", created_at="2026-03-01T10:00:00+09:00"),
        _row("1", "b", created_at="2026-02-01T23:30:00+09:00"),
        _row("1", "c", created_at=""),
        _row("2", "a", name="배"),
    )
    summaries = {s.name: s for s in product_summaries(frame)}
    assert summaries["사과"].created_at == "2026-02-01T23:30:00+09:00" and summaries["배"].created_at == ""
    with pytest.raises(dataclasses.FrozenInstanceError):
        summaries["사과"].created_at = "x"  # type: ignore[misc]
    table = summary_table(list(summaries.values()), set())
    assert dict(zip(table["상품명"], table[COL_CREATED], strict=True)) == {"사과": "2026-02-01 23:30", "배": ""}


def test_history_ignores_created_at_only_difference():
    old = to_csv_text(_frame(_row("1")))
    now = to_csv_text(_frame(_row("1", created_at=OLD)))
    assert diff_dictionary(old, now, SETTINGS).empty


def test_backfill_fills_empty_only_and_keeps_updated_columns():
    frame = _frame(_row("1", created_at=""), _row("2", created_at="2025-01-01T00:00:00+09:00"),
                   _row("3", created_at="", updated_at=""))
    new, filled = backfilled_frame(frame)
    assert filled == 1
    assert list(new["created_at"]) == [OLD, "2025-01-01T00:00:00+09:00", ""]
    assert new[["updated_at", "updated_by"]].equals(frame[["updated_at", "updated_by"]])
    assert frame.iloc[0]["created_at"] == ""  # input not mutated
    legacy, count = backfilled_frame(frame.drop(columns=["created_at"]))
    assert count == 2 and legacy.iloc[0]["created_at"] == OLD
