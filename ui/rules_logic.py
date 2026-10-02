"""Pure helpers of the 고급 설정 rule grids (no streamlit): load, filter, fold edits, judge, build CSV.

The tables keep the ORIGINAL Korean headers of the CSV files (+ META columns) as text, so a save
without edits reproduces the file byte for byte. Row id (`_rid`) = position in the loaded file;
rows added in the grid get the next free ids. Rows of other channels are kept but never shown.
"""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

import pandas as pd

from core.dictionary import _parse_bool
from store.base import DataStore, StoreError
from store.csv_codec import from_csv_text, to_csv_text
from store.rules_repo import DEFAULT_CHANNEL, OPTION_RULES_FILE, PRODUCT_ROUTE_FILE
from store.rules_validators import (
    RULE_COLUMN_LABELS,
    build_config,
    final_guard,
    int_value,
    validate_option_rules,
    validate_product_route,
)
from store.validators import Issue

RID, DELETE = "_rid", "_delete"
CHANNEL, ENABLED, UPDATED_AT, UPDATED_BY = "channel", "enabled", "updated_at", "updated_by"
V_ENABLED, V_PRIORITY, V_KEYWORD, V_VENDOR = "사용", "우선순위", "키워드", "양식명칭"
V_ORDER, V_TARGET, V_ACTION, V_PARAM, V_DESC, V_DELETE = "순서", "적용대상", "ActionType", "Parameter", "설명", "삭제"
ALL_FILTER = "전체"
KIND_BOOL, KIND_INT, KIND_TEXT = "bool", "int", "text"


@dataclass(frozen=True)
class ViewColumn:
    view: str
    header: str  # header in the CSV file
    kind: str


@dataclass(frozen=True)
class TableSpec:
    kind: str
    file: str
    label: str
    number_header: str  # the sort key column (우선순위 / 순서)
    columns: tuple[ViewColumn, ...]
    canon: dict[str, str]  # canonical validator column -> file header
    has_delete: bool
    renumber: bool


ROUTE_SPEC = TableSpec(
    kind="route", file=PRODUCT_ROUTE_FILE, label="상품분류", number_header="우선순위",
    columns=(
        ViewColumn(V_ENABLED, ENABLED, KIND_BOOL), ViewColumn(V_PRIORITY, "우선순위", KIND_INT),
        ViewColumn(V_KEYWORD, "키워드", KIND_TEXT), ViewColumn(V_VENDOR, "양식명칭", KIND_TEXT),
    ),
    canon={"priority": "우선순위", "keyword": "키워드", "vendor": "양식명칭"}, has_delete=False, renumber=False,
)
OPTIONS_SPEC = TableSpec(
    kind="options", file=OPTION_RULES_FILE, label="옵션규칙", number_header="순서",
    columns=(
        ViewColumn(V_ENABLED, ENABLED, KIND_BOOL), ViewColumn(V_ORDER, "순서", KIND_INT),
        ViewColumn(V_VENDOR, "양식명칭", KIND_TEXT), ViewColumn(V_TARGET, "적용대상(상품명)", KIND_TEXT),
        ViewColumn(V_ACTION, "ActionType (명령)", KIND_TEXT), ViewColumn(V_PARAM, "Parameter (설정값)", KIND_TEXT),
        ViewColumn(V_DESC, "설명 (비고)", KIND_TEXT),
    ),
    canon={
        "order": "순서", "vendor": "양식명칭", "target": "적용대상(상품명)", "action": "ActionType (명령)",
        "param": "Parameter (설정값)",
    },
    has_delete=True, renumber=True,
)


@dataclass(frozen=True)
class TableState:
    """One rule file as loaded: `base` (file columns, text), `work` (base + edits + `_delete`), sha, raw text."""

    spec: TableSpec
    base: pd.DataFrame
    work: pd.DataFrame
    sha: str | None
    text: str


@dataclass(frozen=True)
class Judgement:
    errors: list[str]
    warnings: list[str]
    reference: list[str]  # problems of rows that were not touched (never block)
    new_text: str
    modified: list[int]
    added: list[int]
    deleted: list[int]

    @property
    def has_changes(self) -> bool:
        return bool(self.modified or self.added or self.deleted)


# ---------------------------------------------------------------- loading

def _s(value: object) -> str:
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()


def read_table(text: str, spec: TableSpec) -> pd.DataFrame:
    """The CSV as text cells; a missing META column is added, a missing rule column is an error."""
    frame = from_csv_text(text, as_text=True)
    missing = [c.header for c in spec.columns if c.header not in frame.columns and c.header != ENABLED]
    if missing:
        raise ValueError(f"{spec.file}에 필요한 칸이 없습니다: {', '.join(missing)}")
    defaults = {CHANNEL: DEFAULT_CHANNEL, ENABLED: "1", UPDATED_AT: "", UPDATED_BY: ""}
    for name, value in defaults.items():
        if name not in frame.columns:
            frame[name] = value
    return frame.reset_index(drop=True)


def make_work(base: pd.DataFrame) -> pd.DataFrame:
    work = base.copy()
    work[DELETE] = False
    return work


def load_table(store: DataStore, spec: TableSpec) -> TableState:
    snapshot = store.read_text(spec.file)
    if snapshot is None:
        raise StoreError(f"데이터 저장소에 규칙 파일 '{spec.file}'이(가) 없습니다.")
    base = read_table(snapshot.content, spec)
    return TableState(spec, base, make_work(base), snapshot.sha, snapshot.content)


def with_work(state: TableState, work: pd.DataFrame) -> TableState:
    return TableState(state.spec, state.base, work, state.sha, state.text)


# ---------------------------------------------------------------- view

def _number_key(work: pd.DataFrame, spec: TableSpec) -> pd.Series:
    return pd.to_numeric(work[spec.number_header], errors="coerce")


def visible_rids(work: pd.DataFrame, spec: TableSpec) -> list[int]:
    """Rows of the data channel (deleted rows only stay visible where the grid has a 삭제 column)."""
    mask = work[CHANNEL] == DEFAULT_CHANNEL
    if not spec.has_delete:
        mask &= ~work[DELETE]
    return [int(i) for i in work.index[mask]]


def sort_rids(work: pd.DataFrame, rids: Sequence[int], spec: TableSpec) -> list[int]:
    """Stable order by 우선순위 / 순서 (unreadable numbers last)."""
    keys = _number_key(work, spec).loc[list(rids)]
    return [int(i) for i in keys.sort_values(kind="stable", na_position="last").index]


def filter_option_rids(work: pd.DataFrame, vendor: str, action: str, query: str) -> list[int]:
    """Option rows matching the filters (양식명칭 / ActionType / text in 적용대상, Parameter, 설명)."""
    rids = sort_rids(work, visible_rids(work, OPTIONS_SPEC), OPTIONS_SPEC)
    part = work.loc[rids]
    mask = pd.Series(True, index=part.index)
    if vendor != ALL_FILTER:
        mask &= part["양식명칭"].str.strip() == vendor
    if action != ALL_FILTER:
        mask &= part["ActionType (명령)"].str.strip().str.upper() == action
    text = query.strip().lower()
    if text:
        haystack = (part["적용대상(상품명)"] + " " + part["Parameter (설정값)"] + " " + part["설명 (비고)"]).str.lower()
        mask &= haystack.str.contains(text, regex=False)
    return [int(i) for i in part.index[mask]]


def _cell_value(kind: str, text: str) -> object:
    if kind == KIND_BOOL:
        return _parse_bool(text)
    if kind == KIND_INT:
        return int_value(text)
    return text


def build_view(work: pd.DataFrame, rids: Sequence[int], spec: TableSpec) -> pd.DataFrame:
    """Rows shown in the editor (plain 0..n index, `_rid` column first); integer columns hold None where the text is no integer."""
    part = work.loc[list(rids)]
    view = pd.DataFrame({RID: [int(i) for i in part.index]})
    for col in spec.columns:
        values = [_cell_value(col.kind, _s(v)) for v in part[col.header]]
        view[col.view] = pd.Series(values, dtype="Int64") if col.kind == KIND_INT else pd.Series(values, dtype=object)
    if spec.has_delete:
        view[V_DELETE] = [bool(v) for v in part[DELETE]]
    return view


# ---------------------------------------------------------------- folding edits

def _norm(kind: str, value: object) -> object:
    if value is None or value is pd.NA or (isinstance(value, float) and pd.isna(value)):
        return False if kind == KIND_BOOL else ""
    if kind == KIND_BOOL:
        return bool(value) if not isinstance(value, str) else _parse_bool(value)
    if kind == KIND_INT:
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()
    return str(value).strip()


def _to_text(kind: str, value: object) -> str:
    normal = _norm(kind, value)
    return ("1" if normal else "0") if kind == KIND_BOOL else str(normal)


def _blank_row(edited_row: pd.Series, spec: TableSpec) -> bool:
    return all(_norm(c.kind, edited_row.get(c.view)) in ("", False) for c in spec.columns if c.header != ENABLED)


def _new_row(columns: Sequence[str], edited_row: pd.Series, spec: TableSpec, number: str) -> dict[str, object]:
    row: dict[str, object] = dict.fromkeys(columns, "")
    row[CHANNEL] = DEFAULT_CHANNEL
    for col in spec.columns:
        row[col.header] = _to_text(col.kind, edited_row.get(col.view))
    if not row[spec.number_header] and spec.renumber:
        row[spec.number_header] = number
    row[ENABLED] = _to_text(KIND_BOOL, edited_row.get(V_ENABLED, True))
    row[DELETE] = bool(_norm(KIND_BOOL, edited_row.get(V_DELETE))) if spec.has_delete else False
    return row


def fold_grid(
    work: pd.DataFrame, view0: pd.DataFrame, edited: pd.DataFrame, spec: TableSpec, allow_remove: bool
) -> pd.DataFrame:
    """New working copy: the editor's result folded onto `work` (the copy the editor was opened with).

    Only cells that differ from what the editor was given (`view0`) are applied, rows are matched
    by `_rid`. With `allow_remove`, rows the editor dropped are marked deleted; rows without an
    id (added in the grid) get new ids. Calling it again with the same inputs gives the same result.
    """
    updated = work.copy()
    given = view0.set_index(RID)
    seen: set[int] = set()
    additions: list[dict[str, object]] = []
    top = _number_key(work, spec).max(skipna=True)
    next_number = (0 if pd.isna(top) else int(top)) + 1
    for _, row in edited.iterrows():
        raw_rid = row.get(RID)
        if raw_rid is None or pd.isna(raw_rid):
            if not _blank_row(row, spec):
                additions.append(_new_row(list(work.columns), row, spec, str(next_number + len(additions))))
            continue
        rid = int(raw_rid)
        if rid not in updated.index or rid not in given.index:
            continue
        seen.add(rid)
        for col in spec.columns:
            if _norm(col.kind, row[col.view]) != _norm(col.kind, given.at[rid, col.view]):
                updated.at[rid, col.header] = _to_text(col.kind, row[col.view])
        if spec.has_delete:
            updated.at[rid, DELETE] = bool(row[V_DELETE])
    if allow_remove:
        for rid in given.index:
            if int(rid) not in seen:
                updated.at[int(rid), DELETE] = True
    if additions:
        start = int(updated.index.max()) + 1 if len(updated) else 0
        new = pd.DataFrame(additions, index=range(start, start + len(additions)), columns=list(work.columns))
        new[DELETE] = new[DELETE].astype(bool)
        updated = pd.concat([updated, new])
        updated[DELETE] = updated[DELETE].astype(bool)
    return updated


# ---------------------------------------------------------------- changes and saving

def _content_columns(base: pd.DataFrame) -> list[str]:
    return [c for c in base.columns if c not in (UPDATED_AT, UPDATED_BY)]


def changed_rids(base: pd.DataFrame, work: pd.DataFrame) -> tuple[list[int], list[int], list[int]]:
    """(modified, added, deleted) row ids; a deleted row counts only as deleted, added+deleted as nothing."""
    columns = _content_columns(base)
    modified: list[int] = []
    added: list[int] = []
    deleted: list[int] = []
    for rid in work.index:
        rid = int(rid)
        in_base = rid in base.index
        if work.at[rid, DELETE]:
            if in_base:
                deleted.append(rid)
        elif not in_base:
            added.append(rid)
        elif any(_s(base.at[rid, c]) != _s(work.at[rid, c]) for c in columns):
            modified.append(rid)
    return modified, added, deleted


def _sorted_for_save(kept: pd.DataFrame, spec: TableSpec) -> pd.DataFrame:
    """Stable sort by 우선순위 / 순서 (ties keep their position); options are renumbered 1..n."""
    keys = _number_key(kept, spec)
    ordered = kept.loc[keys.sort_values(kind="stable", na_position="last").index].copy()
    if spec.renumber:
        ordered[spec.number_header] = [str(i) for i in range(1, len(ordered) + 1)]
    return ordered


def build_csv(base: pd.DataFrame, work: pd.DataFrame, spec: TableSpec, user: str, now: str) -> str:
    """CSV text to write: deleted rows dropped, changed/added rows stamped, then sorted (and renumbered)."""
    modified, added, _deleted = changed_rids(base, work)
    out = work.copy()
    stamped = [*modified, *added]
    out.loc[stamped, UPDATED_AT] = now
    out.loc[stamped, UPDATED_BY] = user
    kept = _sorted_for_save(out[~out[DELETE]], spec)
    return to_csv_text(kept[list(base.columns)].reset_index(drop=True))


def canonical_frame(work: pd.DataFrame, spec: TableSpec) -> pd.DataFrame:
    """Rows to validate (data channel, not deleted) with the validators' column names; index = row id."""
    kept = work[(~work[DELETE]) & (work[CHANNEL] == DEFAULT_CHANNEL)]
    canon = pd.DataFrame({name: kept[header] for name, header in spec.canon.items()})
    canon[ENABLED] = kept[ENABLED]
    return canon


def describe_issue(canon: pd.DataFrame, issue: Issue, spec: TableSpec) -> str:
    """One Korean line: which rule (keyword / order + action), which column, then the message."""
    if issue.row is None:
        return issue.message
    row = canon.iloc[issue.row]
    who = f"[{_s(row['keyword'])[:30]}]" if spec.kind == "route" else f"[순서 {_s(row['order'])} · {_s(row['action'])}]"
    return f"{who} {RULE_COLUMN_LABELS.get(issue.column, issue.column)}: {issue.message}"


def _unique(lines: list[str]) -> list[str]:
    return list(dict.fromkeys(lines))


def judge(
    state: TableState, work: pd.DataFrame, vendor_ids: Sequence[str], other_csv: str,
    layout: pd.DataFrame, user: str, now: str,
) -> Judgement:
    """Validate an edit. Rows touched by the edit are judged (errors block, warnings need a confirmation);
    problems of untouched rows are reference only. DEFAULT and the final guard always apply."""
    spec = state.spec
    modified, added, deleted = changed_rids(state.base, work)
    new_text = build_csv(state.base, work, spec, user, now)
    canon = canonical_frame(work, spec)
    issues = validate_product_route(canon, vendor_ids) if spec.kind == "route" else validate_option_rules(canon, vendor_ids)
    issues += final_guard(*((new_text, other_csv) if spec.kind == "route" else (other_csv, new_text)), layout)
    touched = {*modified, *added}
    errors: list[str] = []
    warnings: list[str] = []
    reference: list[str] = []
    for issue in issues:
        line = describe_issue(canon, issue, spec)
        if issue.row is None or int(canon.index[issue.row]) in touched:
            (errors if issue.level == "error" else warnings).append(line)
        else:
            reference.append(line)
    return Judgement(_unique(errors), _unique(warnings), _unique(reference), new_text, modified, added, deleted)


def summary_message(spec: TableSpec, judgement: Judgement) -> str:
    return (
        f"{spec.label} 수정: 수정 {len(judgement.modified)}, 추가 {len(judgement.added)}, "
        f"삭제 {len(judgement.deleted)}"
    )


def edited_config(
    route: TableState, options: TableState, layout: pd.DataFrame, user: str, now: str
) -> dict:
    """The config the pipeline would use if both edited tables were saved (raises ValueError and friends)."""
    return build_config(
        build_csv(route.base, route.work, ROUTE_SPEC, user, now),
        build_csv(options.base, options.work, OPTIONS_SPEC, user, now),
        layout,
    )
