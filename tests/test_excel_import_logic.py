"""Excel import: format errors, planning by key, round trip, validation and the extra rule-page logic."""
import io

import pandas as pd
import pytest

from core.dictionary import DictionarySettings
from store.rules_repo import DICTIONARY_COLUMNS
from ui.dictionary_logic import EXPORT_HEADERS, to_excel_bytes
from ui.excel_import_logic import (
    ImportFormatError,
    import_message,
    judge_import,
    plan_import,
    read_import_frame,
)

SETTINGS = DictionarySettings()
NOW, USER = "2026-02-02T10:00:00+09:00", "u@example.com"
IDS = ["V1", "V2"]


def row(no: str, name: str, **kw: str) -> dict[str, str]:
    base = dict.fromkeys(DICTIONARY_COLUMNS, "")
    return {**base, "channel": "naver", "enabled": "1", "needs_review": "0", "append_to_end": "0", "source": "manual",
            "product_no": no, "option_key": "옵션: A", "product_name_ref": name, "option_raw_ref": "옵션: A",
            "vendor_id": "V1", "display_template": f"{name} {{수량}}개", "updated_at": "old", "updated_by": "old", **kw}


def frame(*rows: dict[str, str]) -> pd.DataFrame:
    return pd.DataFrame(list(rows), columns=DICTIONARY_COLUMNS)


def xlsx(df: pd.DataFrame, rename: dict[str, str] | None = None, drop: list[str] | None = None) -> bytes:
    out = df.rename(columns=EXPORT_HEADERS).rename(columns=rename or {}).drop(columns=drop or [])
    buffer = io.BytesIO()
    out.to_excel(buffer, index=False)
    return buffer.getvalue()


CURRENT = frame(row("1", "사과"), row("2", "배", unit_weight_kg="0.5", sum_group="배", display_template=""),
                row("3", "감", enabled="0"))


def test_round_trip_export_import_has_no_changes():
    incoming = read_import_frame(to_excel_bytes(CURRENT))
    assert incoming.shape == CURRENT.shape
    plan = plan_import(CURRENT, incoming, SETTINGS, USER, NOW)
    assert (plan.added, plan.modified, plan.deleted, plan.unchanged) == (0, 0, 0, 3)
    assert plan.changes.empty and plan.frame.equals(CURRENT.fillna("").astype(str))


def test_missing_and_unknown_columns_are_clear_errors():
    with pytest.raises(ImportFormatError, match="빠진 열: 발주양식"):
        read_import_frame(xlsx(CURRENT, drop=["발주양식"]))
    with pytest.raises(ImportFormatError, match="알 수 없는 열: 이상한열"):
        read_import_frame(xlsx(CURRENT, rename={"발주양식": "이상한열"}))
    with pytest.raises(ImportFormatError, match="읽을 수 없습니다"):
        read_import_frame(b"not an excel file")


def test_added_modified_unchanged_and_stamps():
    edited = CURRENT.copy()
    edited.loc[0, "display_template"] = "사과 {수량}박스"
    edited.loc[0, "needs_review"] = "1"
    edited = pd.concat([edited, frame(row("9", "귤", source="", updated_at="", updated_by="", needs_review="1"))], ignore_index=True)
    plan = plan_import(CURRENT, read_import_frame(xlsx(edited)), SETTINGS, USER, NOW)
    assert (plan.added, plan.modified, plan.unchanged, plan.deleted) == (1, 1, 2, 0)
    assert plan.changes["구분"].tolist() == ["수정", "추가"]
    assert "발주서 표기: 사과 {수량}개 → 사과 {수량}박스" in plan.changes.iloc[0]["변경 내용"]
    out = plan.frame
    assert out.loc[0, ["updated_at", "updated_by"]].tolist() == [NOW, USER] and out.loc[0, "needs_review"] == "1"
    assert out.loc[1, "updated_at"] == "old"  # untouched row keeps its stamp
    new = out.iloc[-1]
    assert new["source"] == "excel" and new["updated_by"] == USER and new["needs_review"] == "1" and new["enabled"] == "1"
    assert plan.touched == [0, 3]
    assert import_message(plan) == "엑셀 가져오기: 추가 1, 수정 1, 삭제 0"


def test_rows_missing_from_file_are_kept_unless_delete_requested():
    incoming = read_import_frame(xlsx(CURRENT.iloc[:1]))
    keep = plan_import(CURRENT, incoming, SETTINGS, USER, NOW, delete_missing=False)
    assert len(keep.frame) == 3 and keep.deleted == 0
    drop = plan_import(CURRENT, incoming, SETTINGS, USER, NOW, delete_missing=True)
    assert len(drop.frame) == 1 and drop.deleted == 2
    assert drop.changes["구분"].tolist() == ["삭제", "삭제"] and import_message(drop).endswith("삭제 2")


def test_matching_uses_normalized_key_and_duplicates_are_errors():
    incoming = frame(row("1", "사과", option_key="옵션: A"), row("1", "사과", option_key="옵션: A"))
    plan = plan_import(CURRENT, read_import_frame(xlsx(incoming)), SETTINGS, USER, NOW)
    assert plan.errors and "중복" in plan.errors[0]


def test_validation_blocks_only_touched_rows():
    broken_old = CURRENT.copy()
    broken_old.loc[2, "vendor_id"] = "없는양식"  # untouched row with a problem -> reference only
    edited = broken_old.copy()
    edited.loc[0, "vendor_id"] = "없는양식2"
    plan = plan_import(broken_old, read_import_frame(xlsx(edited)), SETTINGS, USER, NOW)
    judged = judge_import(plan, IDS, SETTINGS)
    assert any("없는양식2" in e for e in judged.errors) and not any("없는양식'" in e for e in judged.errors)
    assert judged.reference == [] or all("감" in r for r in judged.reference)
    clean = judge_import(plan_import(CURRENT, read_import_frame(xlsx(CURRENT)), SETTINGS, USER, NOW), IDS, SETTINGS)
    assert clean.errors == []
