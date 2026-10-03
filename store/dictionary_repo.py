"""Item dictionary and its settings, stored in the data repository."""
from __future__ import annotations

import math
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone

import pandas as pd

from core.dictionary import DictionarySettings, ItemDictionary
from core.option_key import make_option_key, normalize_product_no
from store.base import Author, ConflictError, DataStore, Snapshot
from store.csv_codec import BOM, from_csv_text, to_csv_text
from store.rules_repo import DICTIONARY_COLUMNS, DICTIONARY_FILE, SETTINGS_FILE

KST = timezone(timedelta(hours=9))


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


def load_state(store: DataStore) -> tuple[ItemDictionary, DictionarySettings, str | None]:
    """Dictionary, settings and the dictionary's sha in one read (missing dictionary -> empty, None)."""
    settings = load_settings(store)
    snapshot = store.read_text(DICTIONARY_FILE)
    if snapshot is None:
        return ItemDictionary.empty(), settings, None
    if not snapshot.content.strip(BOM + " \r\n"):
        return ItemDictionary.empty(), settings, snapshot.sha
    return ItemDictionary(from_csv_text(snapshot.content, as_text=True), settings), settings, snapshot.sha


def empty_frame() -> pd.DataFrame:
    return pd.DataFrame(columns=DICTIONARY_COLUMNS)


def frame_from_snapshot(snapshot: Snapshot | None) -> pd.DataFrame:
    """Raw dictionary rows as text, reindexed to DICTIONARY_COLUMNS (missing file -> empty frame)."""
    if snapshot is None or not snapshot.content.strip(BOM + " \r\n"):
        return empty_frame()
    frame = from_csv_text(snapshot.content, as_text=True)
    return frame.reindex(columns=DICTIONARY_COLUMNS, fill_value="")


def load_dictionary_frame(store: DataStore) -> tuple[pd.DataFrame, str | None]:
    snapshot = store.read_text(DICTIONARY_FILE)
    return frame_from_snapshot(snapshot), (snapshot.sha if snapshot else None)


def save_dictionary(
    store: DataStore, frame: pd.DataFrame, expected_sha: str | None, author: Author, message: str
) -> str:
    """Write the whole dictionary. ConflictError propagates when the file changed since `expected_sha`."""
    text = to_csv_text(frame.reindex(columns=DICTIONARY_COLUMNS, fill_value=""))
    return store.write_text(DICTIONARY_FILE, text, expected_sha, message, author)


def now_kst_iso() -> str:
    return datetime.now(KST).isoformat(timespec="seconds")


def entry_key(product_no: object, option_key: object, settings: DictionarySettings) -> tuple[str, str]:
    return (
        normalize_product_no(product_no),
        make_option_key(option_key, settings.ignored_option_groups),
    )


def _is_enabled(value: object) -> bool:
    return str(value).strip().lower() in {"1", "true", "t", "y", "yes", "예", "o"}


def _enabled_keys(frame: pd.DataFrame, settings: DictionarySettings) -> set[tuple[str, str]]:
    return {
        entry_key(row["product_no"], row["option_key"], settings)
        for _, row in frame.iterrows()
        if _is_enabled(row["enabled"])
    }


def _new_row(row: dict, author: Author, now: str, settings: DictionarySettings) -> dict[str, str]:
    base = dict.fromkeys(DICTIONARY_COLUMNS, "")
    for column in DICTIONARY_COLUMNS:
        if column in row and row[column] is not None and not _is_nan(row[column]):
            base[column] = str(row[column])
    product_no, option_key = entry_key(row.get("product_no"), row.get("option_key"), settings)
    base.update(
        product_no=product_no, option_key=option_key, channel="naver", enabled="1", needs_review="1",
        source="manual", updated_at=now, updated_by=author.name, last_seen_at="",
    )
    if not base["append_to_end"]:
        base["append_to_end"] = "0"
    return base


def _is_nan(value: object) -> bool:
    return isinstance(value, float) and math.isnan(value)


@dataclass(frozen=True)
class AppendResult:
    added: list[tuple[str, str]]
    skipped_existing: list[tuple[str, str]]
    sha: str


def _merge_rows(
    frame: pd.DataFrame, rows: list[dict], author: Author, settings: DictionarySettings
) -> tuple[pd.DataFrame, list[tuple[str, str]], list[tuple[str, str]]]:
    present = _enabled_keys(frame, settings)
    now = now_kst_iso()
    added: list[tuple[str, str]] = []
    skipped: list[tuple[str, str]] = []
    new_rows: list[dict[str, str]] = []
    for row in rows:
        key = entry_key(row.get("product_no"), row.get("option_key"), settings)
        if key in present:
            skipped.append(key)
            continue
        present.add(key)
        added.append(key)
        new_rows.append(_new_row(row, author, now, settings))
    if not new_rows:
        return frame, added, skipped
    merged = pd.concat([frame, pd.DataFrame(new_rows, columns=DICTIONARY_COLUMNS)], ignore_index=True)
    return merged, added, skipped


def append_entries(
    store: DataStore,
    rows: list[dict],
    author: Author,
    settings: DictionarySettings,
    max_retries: int = 3,
) -> AppendResult:
    """Add new dictionary rows (flagged for review). Keys already present are skipped.

    The write uses the sha of the read it was based on; on ConflictError the file is re-read and
    the append is retried up to `max_retries` times, then the error is raised.
    """
    for attempt in range(max_retries + 1):
        snapshot = store.read_text(DICTIONARY_FILE)
        frame = frame_from_snapshot(snapshot)
        sha = snapshot.sha if snapshot else None
        merged, added, skipped = _merge_rows(frame, rows, author, settings)
        if not added:
            return AppendResult([], skipped, sha or "")
        message = f"사전 등록 {len(added)}건 (검토 필요)"
        try:
            new_sha = store.write_text(DICTIONARY_FILE, to_csv_text(merged), sha, message, author)
        except ConflictError:
            if attempt == max_retries:
                raise
            continue
        return AppendResult(added, skipped, new_sha)
    raise ConflictError("unreachable")  # pragma: no cover


def _all_keys(frame: pd.DataFrame, settings: DictionarySettings) -> set[tuple[str, str]]:
    """Keys of every row, enabled or not (a disabled row is a decision too)."""
    return {entry_key(r["product_no"], r["option_key"], settings) for _, r in frame.fillna("").iterrows()}


def add_missing_entries(
    store: DataStore,
    candidates: pd.DataFrame,
    author: Author,
    settings: DictionarySettings,
    max_retries: int = 3,
) -> AppendResult:
    """Add candidate rows (e.g. from the seed tool) whose key is not in the dictionary yet.

    Existing rows - edited, reviewed or disabled - are never touched. Candidate rows keep their
    own fields (source, needs_review, templates). Conflicts are retried like append_entries.
    """
    rows = candidates.reindex(columns=DICTIONARY_COLUMNS, fill_value="").fillna("").astype(str)
    for attempt in range(max_retries + 1):
        snapshot = store.read_text(DICTIONARY_FILE)
        frame = frame_from_snapshot(snapshot)
        sha = snapshot.sha if snapshot else None
        known = _all_keys(frame, settings)
        added: list[tuple[str, str]] = []
        skipped: list[tuple[str, str]] = []
        new_rows = []
        for _, row in rows.iterrows():
            key = entry_key(row["product_no"], row["option_key"], settings)
            if key in known:
                skipped.append(key)
                continue
            known.add(key)
            added.append(key)
            new_rows.append({**row.to_dict(), "product_no": key[0], "option_key": key[1]})
        if not added:
            return AppendResult([], skipped, sha or "")
        merged = pd.concat([frame.reindex(columns=DICTIONARY_COLUMNS), pd.DataFrame(new_rows, columns=DICTIONARY_COLUMNS)],
                           ignore_index=True)
        message = f"사전 초안 추가 {len(added)}건 (기존 항목 유지, 검토 필요)"
        try:
            new_sha = store.write_text(DICTIONARY_FILE, to_csv_text(merged), sha, message, author)
        except ConflictError:
            if attempt == max_retries:
                raise
            continue
        return AppendResult(added, skipped, new_sha)
    raise ConflictError("unreachable")  # pragma: no cover


# ---------------------------------------------------------------- per-product save

RowKey = tuple[str, str]


def _text_frame(frame: pd.DataFrame) -> pd.DataFrame:
    return frame.reindex(columns=DICTIONARY_COLUMNS, fill_value="").fillna("").astype(str).reset_index(drop=True)


def text_records(frame: pd.DataFrame) -> list[dict[str, str]]:
    """Rows as dicts of text, all DICTIONARY_COLUMNS."""
    return [{str(k): str(v) for k, v in r.items()} for r in _text_frame(frame).to_dict("records")]


GROUP_SEP = "\x1f"


def group_id(product_no: str, product_name_ref: str) -> str:
    """Stable id of a product group: main product and add-ons share a product_no but differ by name."""
    return f"{product_no}{GROUP_SEP}{product_name_ref.strip()}"


def group_id_of(row: Mapping[str, str]) -> str:
    return group_id(str(row["product_no"]), str(row["product_name_ref"]))


def group_ids(text: pd.DataFrame) -> pd.Series:
    """group_id of every row of a text frame (same index)."""
    return text["product_no"] + GROUP_SEP + text["product_name_ref"].str.strip()


def product_rows(frame: pd.DataFrame, group: str) -> dict[RowKey, tuple[str, ...]]:
    """All columns (as text) of one group's rows, by (product_no, option_key)."""
    text = _text_frame(frame)
    part = text[group_ids(text) == group]
    return {(str(r["product_no"]), str(r["option_key"])): tuple(r) for _, r in part.iterrows()}


def _replace_product(latest: pd.DataFrame, new_frame: pd.DataFrame, group: str) -> pd.DataFrame:
    """`latest` with this group's rows replaced (or removed) by their versions in `new_frame`."""
    new_rows = {(r["product_no"], r["option_key"]): r for r in text_records(new_frame) if group_id_of(r) == group}
    kept: list[dict[str, str]] = []
    for row in text_records(latest):
        if group_id_of(row) != group:
            kept.append(row)
        elif (key := (row["product_no"], row["option_key"])) in new_rows:
            kept.append(new_rows[key])
    return pd.DataFrame(kept, columns=DICTIONARY_COLUMNS)


def save_product(
    store: DataStore,
    base_frame: pd.DataFrame,
    base_sha: str | None,
    new_frame: pd.DataFrame,
    group: str,
    author: Author,
    message: str,
) -> str:
    """Save `new_frame` (the base with one group edited). Returns the new sha.

    When the file changed meanwhile, the edit is re-applied onto the latest file if nobody touched
    THIS group (other groups keep their newer versions); otherwise ConflictError.
    """
    try:
        return save_dictionary(store, new_frame, base_sha, author, message)
    except ConflictError:
        latest, latest_sha = load_dictionary_frame(store)
        if product_rows(latest, group) != product_rows(base_frame, group):
            raise
        merged = _replace_product(latest, new_frame, group)
        return save_dictionary(store, merged, latest_sha, author, message)
