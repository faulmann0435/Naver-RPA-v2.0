"""Purchase-order forms (OutputLayout) stored as one CSV file in the data repository.

One row = one output column of one form. The CSV keeps the ORIGINAL Korean sheet headers of
config.xlsx plus META columns, so it can be edited as a plain table and converted back losslessly.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

from store.base import DataStore
from store.csv_codec import from_csv_text, to_csv_text

OUTPUT_LAYOUT_FILE = "output_layout.csv"
LAYOUT_CHANNEL = "naver"
FORM, FILENAME, COLUMN, HEADER, SOURCE, FIXED = (
    "양식명칭", "파일명", "열", "헤더명", "매핑데이터", "고정값(Hardcoded)",
)
LAYOUT_HEADERS = [FORM, FILENAME, COLUMN, HEADER, SOURCE, FIXED]
LAYOUT_META = ["channel", "updated_at", "updated_by"]
_INT_TEXT = re.compile(r"^-?(0|[1-9]\d*)$")


@dataclass(frozen=True)
class LayoutSnapshot:
    """The layout of the store: raw sheet-shaped frame, the sha it was read at and its CSV text."""

    raw: pd.DataFrame
    sha: str
    text: str


def _drop_empty_unnamed(sheet: pd.DataFrame) -> pd.DataFrame:
    keep = [c for c in sheet.columns if not (str(c).startswith("Unnamed") and sheet[c].isna().all())]
    return sheet[keep]


def layout_sheet_to_csv(sheet: pd.DataFrame, migrated_by: str, now: str) -> str:
    """OutputLayout sheet (as read from config.xlsx) -> CSV text with META columns."""
    out = _drop_empty_unnamed(sheet).copy()
    out["channel"] = LAYOUT_CHANNEL
    out["updated_at"] = now
    out["updated_by"] = migrated_by
    return to_csv_text(out)


def _excel_like(frame: pd.DataFrame) -> pd.DataFrame:
    """Text cells -> what pd.read_excel gives: empty = NaN, whole numbers in 고정값 = int."""
    out = frame.astype(object).copy()
    for name in out.columns:
        out[name] = [None if v == "" else v for v in out[name]]
        out[name] = out[name].where(out[name].notna(), float("nan"))
    if FIXED in out.columns:
        out[FIXED] = [int(v) if isinstance(v, str) and _INT_TEXT.match(v) else v for v in out[FIXED]]
    return out


def layout_csv_to_raw(text: str) -> pd.DataFrame:
    """CSV text -> the raw OutputLayout frame (channel naver only, META dropped)."""
    frame = from_csv_text(text, as_text=True)
    if "channel" in frame.columns:
        frame = frame[frame["channel"] == LAYOUT_CHANNEL]
    frame = frame.drop(columns=[c for c in LAYOUT_META if c in frame.columns]).reset_index(drop=True)
    return _excel_like(frame)


def load_layout(store: DataStore) -> LayoutSnapshot | None:
    """output_layout.csv of the store, or None when the file does not exist (not migrated yet)."""
    snapshot = store.read_text(OUTPUT_LAYOUT_FILE)
    if snapshot is None:
        return None
    return LayoutSnapshot(layout_csv_to_raw(snapshot.content), snapshot.sha, snapshot.content)
