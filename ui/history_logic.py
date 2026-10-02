"""Pure helpers of the 변경 이력 page (no streamlit): revision lists, diffs against the current file, revert."""
from __future__ import annotations

from collections import Counter
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import datetime, timezone

import pandas as pd

from core.dictionary import DictionarySettings
from store.base import Author, DataStore, Revision, StoreError
from store.csv_codec import BOM, from_csv_text
from store.dictionary_repo import KST, empty_frame, entry_key
from store.rules_repo import (
    DICTIONARY_COLUMNS,
    DICTIONARY_FILE,
    OPTION_RULES_FILE,
    PRODUCT_ROUTE_FILE,
    SETTINGS_FILE,
)
from store.rules_validators import (
    final_guard,
    validate_option_rules,
    validate_product_route,
)
from store.validators import validate_dictionary
from ui.dictionary_logic import describe_issue, field_differences
from ui.rules_extra_logic import SETTING_LABELS, parse_settings, settings_values
from ui.rules_logic import (
    CHANNEL,
    OPTIONS_SPEC,
    ROUTE_SPEC,
    UPDATED_AT,
    UPDATED_BY,
    TableSpec,
    canonical_frame,
    make_work,
    read_table,
)

TARGETS: dict[str, str] = {
    "품목 사전": DICTIONARY_FILE, "상품분류": PRODUCT_ROUTE_FILE, "옵션규칙": OPTION_RULES_FILE, "설정": SETTINGS_FILE,
}
REVIVES, VANISHES = "되돌리면 다시 생김", "되돌리면 사라짐"
DICT_DIFF_COLUMNS = ["구분", "상품명", "옵션", "변경 내용"]
RULE_DIFF_COLUMNS = ["구분", "규칙"]
SETTING_DIFF_COLUMNS = ["항목", "선택한 버전", "현재"]
NO_REVISIONS = "이력이 없습니다."


@dataclass(frozen=True)
class DiffResult:
    table: pd.DataFrame
    note: str = ""  # e.g. "내용은 같고 순서만 다릅니다"

    @property
    def empty(self) -> bool:
        return self.table.empty and not self.note


# ---------------------------------------------------------------- revisions

def kst_text(iso: str) -> str:
    """'2026-01-05T00:30:00Z' -> '2026-01-05 09:30' (KST); unreadable text is returned as is."""
    try:
        moment = datetime.fromisoformat(iso.strip().replace("Z", "+00:00"))
    except ValueError:
        return iso
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=timezone.utc)
    return moment.astimezone(KST).strftime("%Y-%m-%d %H:%M")


def revision_label(rev: Revision) -> str:
    return f"{kst_text(rev.date)} · {rev.author} · {rev.message}"


def history_table(revisions: Sequence[Revision]) -> pd.DataFrame:
    return pd.DataFrame(
        [{"시각": kst_text(r.date), "작업자": r.author, "내용": r.message} for r in revisions],
        columns=["시각", "작업자", "내용"],
    )


def file_label(path: str) -> str:
    return next((label for label, file in TARGETS.items() if file == path), path)


# ---------------------------------------------------------------- diffs

def frame_from_text(text: str) -> pd.DataFrame:
    """dictionary.csv text as a text frame with DICTIONARY_COLUMNS (empty file -> empty frame)."""
    if not text.strip(BOM + " \r\n"):
        return empty_frame()
    return from_csv_text(text, as_text=True).reindex(columns=DICTIONARY_COLUMNS, fill_value="").fillna("")


def _rows_by_key(frame: pd.DataFrame, settings: DictionarySettings) -> dict[tuple[str, str], pd.Series]:
    return {entry_key(r["product_no"], r["option_key"], settings): r for _, r in frame.iterrows()}


def _label(row: pd.Series) -> tuple[str, str]:
    name = str(row["product_name_ref"]).strip() or str(row["product_no"])
    return name, str(row["option_raw_ref"]).strip() or str(row["option_key"]).strip() or "옵션 없음"


def diff_dictionary(old_text: str, current_text: str, settings: DictionarySettings) -> DiffResult:
    """Old version -> current file, by (product_no, option_key): 추가 (only now) / 삭제 (only before) / 변경."""
    old = _rows_by_key(frame_from_text(old_text), settings)
    now = _rows_by_key(frame_from_text(current_text), settings)
    rows: list[list[str]] = []
    for key, row in now.items():
        if key not in old:
            rows.append(["추가", *_label(row), f"이 버전 이후에 추가됨 ({VANISHES})"])
            continue
        changes = field_differences(old[key], row)
        if changes:
            rows.append(["변경", *_label(row), " / ".join(changes)])
    rows += [["삭제", *_label(row), f"이 버전 이후에 삭제됨 ({REVIVES})"] for key, row in old.items() if key not in now]
    return DiffResult(pd.DataFrame(rows, columns=DICT_DIFF_COLUMNS))


def _identity_columns(frame: pd.DataFrame, spec: TableSpec) -> list[str]:
    """Full row content without META columns (and without the auto-renumbered 순서)."""
    skip = {CHANNEL, UPDATED_AT, UPDATED_BY, *([spec.number_header] if spec.renumber else [])}
    return [c for c in frame.columns if c not in skip]


def _describe_rule(row: pd.Series, spec: TableSpec) -> str:
    off = "" if str(row.get("enabled", "1")).strip() in ("1", "1.0", "") else " (꺼짐)"
    if spec.kind == "route":
        return f"[우선순위 {row['우선순위']}] {row['키워드']} → {row['양식명칭']}{off}"
    param = str(row["Parameter (설정값)"])
    shown = param if len(param) <= 60 else param[:60] + "…"
    return f"{row['양식명칭']} / {row['적용대상(상품명)']} / {row['ActionType (명령)']} / {shown}{off}"


def _not_in(lines: list[tuple[str, ...]], other: list[tuple[str, ...]]) -> list[int]:
    remaining = Counter(other)
    left: list[int] = []
    for i, line in enumerate(lines):
        if remaining[line] > 0:
            remaining[line] -= 1
        else:
            left.append(i)
    return left


def diff_rules(old_text: str, current_text: str, spec: TableSpec) -> DiffResult:
    """Rows compared by their full content: only in the old version (되돌리면 다시 생김), only in the
    current file (되돌리면 사라짐). Equal rows in a different order are reported as a note."""
    old, now = read_table(old_text, spec), read_table(current_text, spec)
    columns = [c for c in _identity_columns(now, spec) if c in old.columns]
    old_keys = [tuple(str(v).strip() for v in r) for r in old[columns].itertuples(index=False)]
    now_keys = [tuple(str(v).strip() for v in r) for r in now[columns].itertuples(index=False)]
    rows = [[REVIVES, _describe_rule(old.iloc[i], spec)] for i in _not_in(old_keys, now_keys)]
    rows += [[VANISHES, _describe_rule(now.iloc[i], spec)] for i in _not_in(now_keys, old_keys)]
    note = "내용은 같고 줄 순서만 다릅니다." if not rows and old_keys != now_keys else ""
    return DiffResult(pd.DataFrame(rows, columns=RULE_DIFF_COLUMNS), note)


def _show(value: object) -> str:
    if isinstance(value, list):
        return ", ".join(map(str, value)) or "(비어 있음)"
    return str(value)


def diff_settings(old_text: str, current_text: str) -> DiffResult:
    """Key-level before/after of settings.json."""
    old, now = parse_settings(old_text), parse_settings(current_text)
    old_sep, old_groups = settings_values(old)
    now_sep, now_groups = settings_values(now)
    rows = []
    if old_sep != now_sep:
        rows.append([SETTING_LABELS["item_separator"], f'"{old_sep}"', f'"{now_sep}"'])
    if old_groups != now_groups:
        rows.append([SETTING_LABELS["ignored_option_groups"], _show(old_groups), _show(now_groups)])
    for key in sorted((set(old) | set(now)) - set(SETTING_LABELS)):
        if old.get(key) != now.get(key):
            rows.append([key, _show(old.get(key, "(없음)")), _show(now.get(key, "(없음)"))])
    return DiffResult(pd.DataFrame(rows, columns=SETTING_DIFF_COLUMNS))


def diff_against_current(path: str, old_text: str, current_text: str, settings: DictionarySettings) -> DiffResult:
    if path == DICTIONARY_FILE:
        return diff_dictionary(old_text, current_text, settings)
    if path == PRODUCT_ROUTE_FILE:
        return diff_rules(old_text, current_text, ROUTE_SPEC)
    if path == OPTION_RULES_FILE:
        return diff_rules(old_text, current_text, OPTIONS_SPEC)
    return diff_settings(old_text, current_text)


# ---------------------------------------------------------------- revert

def _canon(text: str, spec: TableSpec) -> pd.DataFrame:
    """A rule CSV in the validators' column names (same path as a normal save)."""
    return canonical_frame(make_work(read_table(text, spec)), spec)


def revert_errors(
    path: str, old_text: str, vendor_ids: Sequence[str], settings: DictionarySettings,
    other_rules_csv: str, layout: pd.DataFrame,
) -> list[str]:
    """Reasons why the old content must not be restored (empty list = safe to revert)."""
    if path == DICTIONARY_FILE:
        frame = frame_from_text(old_text)
        return [describe_issue(frame, i) for i in validate_dictionary(frame, vendor_ids, settings) if i.level == "error"]
    # Rules: the same errors that block a normal save (e.g. exactly one DEFAULT row), then the final guard.
    if path in (PRODUCT_ROUTE_FILE, OPTION_RULES_FILE):
        try:
            _canon(old_text, ROUTE_SPEC if path == PRODUCT_ROUTE_FILE else OPTIONS_SPEC)
        except (ValueError, KeyError) as e:
            return [f"이 버전의 파일 형식을 읽을 수 없습니다. ({e})"]
    if path == PRODUCT_ROUTE_FILE:
        checks = validate_product_route(_canon(old_text, ROUTE_SPEC), vendor_ids)
        return [i.message for i in checks if i.level == "error"] + [
            i.message for i in final_guard(old_text, other_rules_csv, layout)
        ]
    if path == OPTION_RULES_FILE:
        checks = validate_option_rules(_canon(old_text, OPTIONS_SPEC), vendor_ids)
        return [i.message for i in checks if i.level == "error"] + [
            i.message for i in final_guard(other_rules_csv, old_text, layout)
        ]
    try:
        parse_settings(old_text)
    except ValueError as e:
        return [f"설정 파일을 읽을 수 없습니다. ({e})"]
    return []


def revert_message(path: str, rev: Revision) -> str:
    return f"되돌리기: {file_label(path)} → {kst_text(rev.date)} 버전"


def old_content(store: DataStore, path: str, rev: Revision) -> str:
    text = store.read_text_at(path, rev.sha)
    if text is None:
        raise StoreError("이 버전의 내용을 불러올 수 없습니다.")
    return text


def revert_file(store: DataStore, path: str, rev: Revision, expected_sha: str | None, author: Author) -> str:
    """Write the old content as a NEW version (history keeps everything). ConflictError when the file moved on."""
    return store.write_text(path, old_content(store, path, rev), expected_sha, revert_message(path, rev), author)
