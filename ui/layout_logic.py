"""Pure helpers of the 발주서 양식 page (no streamlit): view rows, folding edits, building the CSV, judging.

The table keeps the ORIGINAL Korean headers of output_layout.csv (+ META) as text. The page works
on ONE form at a time: rows of other forms stay byte for byte as they are and in place; the edited
form's rows are replaced by the new block (at the position of its first old row; a new form goes to
the end). 열 is recomputed from the 순서 order as A, B, C ... so the user never types letters.
A save without edits returns the original text unchanged.
"""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

import pandas as pd

from core.exporter import ROW_NUMBER_SOURCE, excel_column_number
from store.csv_codec import from_csv_text, to_csv_text
from store.layout_repo import (
    COLUMN,
    FILENAME,
    FIXED,
    FORM,
    HEADER,
    LAYOUT_CHANNEL,
    LAYOUT_HEADERS,
    LAYOUT_META,
    SOURCE,
)
from store.layout_validators import (
    FORBIDDEN_FILENAME_CHARS,
    LayoutReferences,
    layout_final_guard,
    validate_layout,
)

RID = "_rid"
V_ORDER, V_NAME, V_SOURCE, V_FIXED, V_DELETE = "순서", "칸 이름", "넣을 내용", "고정 글자", "삭제"
CHANNEL, UPDATED_AT, UPDATED_BY = LAYOUT_META
BLANK_LABEL = "(빈칸)"
# Internal source name -> plain Korean label. Every column of the order file survives merging
# (core/merger.py keeps all columns of a group: first value, except 수량 = sum, 배송메세지 = joined,
# 결제일 = earliest); processed_option is built by the pipeline.
SOURCE_LABELS: dict[str, str] = {
    "processed_option": "품목 (정리된 옵션)",
    "수취인명": "받는 분 이름",
    "수취인연락처1": "받는 분 연락처",
    "수취인연락처2": "받는 분 연락처 2",
    "통합배송지": "받는 분 주소",
    "우편번호": "우편번호",
    "배송메세지": "배송 메시지",
    "구매자명": "주문한 분 이름",
    "구매자연락처": "주문한 분 연락처",
    "상품주문번호": "상품주문번호 (묶음이면 첫 번째)",
    "주문번호": "주문번호",
    "상품명": "상품명 (묶음이면 첫 번째)",
    "수량": "수량 (묶음 합계)",
    "결제일": "결제일 (묶음 중 가장 빠른 날)",
    ROW_NUMBER_SOURCE: "순번 (1, 2, 3 … 자동)",
}
KNOWN_SOURCES = frozenset(SOURCE_LABELS)
_LABEL_TO_SOURCE = {label: source for source, label in SOURCE_LABELS.items()}


def source_label(source: str) -> str:
    """Label shown for a stored 매핑데이터 value (an unknown value is shown as it is)."""
    return BLANK_LABEL if not source else SOURCE_LABELS.get(source, source)


def source_value(label: str) -> str:
    """Stored 매핑데이터 value for a label picked in the grid."""
    return "" if label in ("", BLANK_LABEL) else _LABEL_TO_SOURCE.get(label, label)


def source_options(extra_sources: Sequence[str] = ()) -> list[str]:
    """Labels offered in the selectbox; sources already in the data but not in the mapping stay selectable."""
    extras = [s for s in dict.fromkeys(extra_sources) if s and s not in SOURCE_LABELS]
    return [BLANK_LABEL, *SOURCE_LABELS.values(), *extras]


def _s(value: object) -> str:
    return "" if value is None or value is pd.NA or (isinstance(value, float) and pd.isna(value)) else str(value).strip()


def column_letters(number: int) -> str:
    """1 -> A, 26 -> Z, 27 -> AA."""
    letters = ""
    while number > 0:
        number, rest = divmod(number - 1, 26)
        letters = chr(ord("A") + rest) + letters
    return letters


# ---------------------------------------------------------------- loading

def read_layout_table(text: str) -> pd.DataFrame:
    """output_layout.csv as text cells; a missing header or META column is added (empty / naver)."""
    frame = from_csv_text(text, as_text=True)
    for name in LAYOUT_HEADERS:
        if name not in frame.columns:
            frame[name] = ""
    for name in LAYOUT_META:
        if name not in frame.columns:
            frame[name] = LAYOUT_CHANNEL if name == CHANNEL else ""
    return frame.reset_index(drop=True)


def form_positions(table: pd.DataFrame, form: str) -> list[int]:
    """Row positions of a form (naver channel), in file order."""
    return [
        i for i, (name, channel) in enumerate(zip(table[FORM], table[CHANNEL]))
        if channel == LAYOUT_CHANNEL and _s(name) == form
    ]


def form_names(table: pd.DataFrame) -> list[str]:
    """Forms of the table in first-appearance order."""
    mask = table[CHANNEL] == LAYOUT_CHANNEL
    return [n for n in dict.fromkeys(_s(v) for v in table.loc[mask, FORM]) if n]


def form_filename(table: pd.DataFrame, form: str) -> str:
    positions = form_positions(table, form)
    return _s(table.at[positions[0], FILENAME]) if positions else ""


def used_sources(table: pd.DataFrame) -> list[str]:
    return [s for s in dict.fromkeys(_s(v) for v in table[SOURCE]) if s]


# ---------------------------------------------------------------- view and grid

def column_order(table: pd.DataFrame, positions: Sequence[int]) -> list[int]:
    """Positions of a form ordered by the column letter (stable; unreadable letters last)."""
    keys = {p: excel_column_number(_s(table.at[p, COLUMN])) for p in positions}
    return sorted(positions, key=lambda p: (keys[p] is None, keys[p] or 0))


def build_view(table: pd.DataFrame, form: str) -> pd.DataFrame:
    """The rows shown in the editor: `_rid` (row position, hidden), 순서 1..n, 칸 이름, 넣을 내용, 고정 글자, 삭제."""
    ordered = column_order(table, form_positions(table, form))
    return pd.DataFrame({
        RID: pd.Series(ordered, dtype="Int64"),
        V_ORDER: pd.Series(range(1, len(ordered) + 1), dtype="Int64"),
        V_NAME: pd.Series([_s(table.at[p, HEADER]) for p in ordered], dtype=object),
        V_SOURCE: pd.Series([source_label(_s(table.at[p, SOURCE])) for p in ordered], dtype=object),
        V_FIXED: pd.Series([_s(table.at[p, FIXED]) for p in ordered], dtype=object),
        V_DELETE: pd.Series([False] * len(ordered), dtype=bool),
    })


@dataclass(frozen=True)
class ColumnRecord:
    """One output column of a form as stored (used to compare two versions of the file)."""

    index: int  # 1-based left-to-right position within the form
    source: str
    fixed: str
    filename: str


def column_records(table: pd.DataFrame) -> dict[tuple[str, str, int], ColumnRecord]:
    """(form, header, n-th time this header appears in the form) -> record, for every form of the table."""
    records: dict[tuple[str, str, int], ColumnRecord] = {}
    for form in form_names(table):
        seen: dict[str, int] = {}
        for index, p in enumerate(column_order(table, form_positions(table, form)), start=1):
            header = _s(table.at[p, HEADER])
            seen[header] = seen.get(header, 0) + 1
            records[(form, header, seen[header])] = ColumnRecord(
                index, _s(table.at[p, SOURCE]), _s(table.at[p, FIXED]), _s(table.at[p, FILENAME])
            )
    return records


@dataclass(frozen=True)
class GridRow:
    rid: int | None  # row position in the loaded table; None = added in the grid
    order: int | None
    name: str
    source: str  # stored value, not the label
    fixed: str
    delete: bool


def _bool(value: object) -> bool:
    return False if value is None or value is pd.NA or (isinstance(value, float) and pd.isna(value)) else bool(value)


def _order_number(value: object) -> int | None:
    number = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    return None if pd.isna(number) else int(number)


def grid_rows(grid: pd.DataFrame) -> list[GridRow]:
    """Editor rows as plain records; rows that were added and left blank are dropped."""
    rows: list[GridRow] = []
    for _, series in grid.iterrows():
        raw_rid = series.get(RID)
        rid = None if raw_rid is None or raw_rid is pd.NA or pd.isna(raw_rid) else int(raw_rid)
        row = GridRow(
            rid, _order_number(series.get(V_ORDER)), _s(series.get(V_NAME)),
            source_value(_s(series.get(V_SOURCE))), _s(series.get(V_FIXED)), _bool(series.get(V_DELETE)),
        )
        if rid is None and not (row.name or row.source or row.fixed):
            continue
        rows.append(row)
    return rows


def final_order(rows: Sequence[GridRow]) -> list[GridRow]:
    """Surviving rows in their new left-to-right order: by 순서, ties (and empty 순서) keep grid order."""
    kept = [r for r in rows if not r.delete]
    indexed = sorted(enumerate(kept), key=lambda t: (t[1].order is None, t[1].order or 0, t[0]))
    return [row for _, row in indexed]


NEW_COLUMN_LABEL = "➕ 새 칸 추가"


def fixed_options(table: pd.DataFrame, grid: pd.DataFrame) -> list[str]:
    """Fixed texts offered for picking: "" first, then every one used in any form or in the grid (most used first)."""
    values = [_s(v) for v in table[FIXED]] + [_s(v) for v in grid[V_FIXED]]
    counts: dict[str, int] = {}
    for value in values:
        if value:
            counts[value] = counts.get(value, 0) + 1
    return ["", *sorted(counts, key=lambda v: -counts[v])]


def column_choices(grid: pd.DataFrame) -> list[str]:
    """Rows of the grid offered for editing one at a time ("3. 받는분성명"), plus adding a new one first."""
    labels = [f"{i}. {_s(name) or '(이름 없음)'}" for i, name in enumerate(grid[V_NAME], start=1)]
    return [NEW_COLUMN_LABEL, *labels]


def set_one_column(grid: pd.DataFrame, index: int | None, name: str, source_label: str, fixed: str) -> pd.DataFrame:
    """A copy of the grid with row `index` (0-based) changed, or a new row appended at the end when None.

    Used by the input box under the grid, where Korean can be typed (the grid's own cells break IME input).
    """
    if index is None:
        orders = [o for o in (_order_number(v) for v in grid[V_ORDER]) if o is not None]
        new_row = pd.DataFrame({
            RID: pd.Series([pd.NA], dtype="Int64"),
            V_ORDER: pd.Series([max(orders, default=0) + 1], dtype="Int64"),
            V_NAME: pd.Series([name.strip()], dtype=object),
            V_SOURCE: pd.Series([source_label], dtype=object),
            V_FIXED: pd.Series([fixed.strip()], dtype=object),
            V_DELETE: pd.Series([False], dtype=bool),
        })
        return pd.concat([grid, new_row], ignore_index=True)
    changed = grid.copy()
    changed.at[index, V_NAME] = name.strip()
    changed.at[index, V_SOURCE] = source_label
    changed.at[index, V_FIXED] = fixed.strip()
    return changed


def preview_table(grid: pd.DataFrame) -> pd.DataFrame:
    """Header row in final order + one example row (fixed text / "(label)" / blank). Empty headers are not exported."""
    headers: list[str] = []
    example: list[str] = []
    for row in final_order(grid_rows(grid)):
        if not row.name:
            continue
        name, n = row.name, 2
        while name in headers:
            name, n = f"{row.name} ({n})", n + 1
        headers.append(name)
        example.append(row.fixed or (f"({source_label(row.source)})" if row.source else ""))
    return pd.DataFrame([example], columns=headers)


# ---------------------------------------------------------------- applying an edit

@dataclass(frozen=True)
class EditResult:
    table: pd.DataFrame
    text: str
    modified: int
    added: int
    deleted: int
    form_positions: tuple[int, ...]  # rows of the edited form in the new table
    changed_positions: tuple[int, ...]  # of those, the rows that are new or changed

    @property
    def has_changes(self) -> bool:
        return bool(self.modified or self.added or self.deleted)


def _records(rows: Sequence[GridRow]) -> list[tuple]:
    return [(r.rid, r.order, r.name, r.source, r.fixed, r.delete) for r in rows]


def _block(
    table: pd.DataFrame, old: set[int], ordered: Sequence[GridRow], form: str, filename: str, user: str, now: str
) -> tuple[list[dict[str, str]], list[bool]]:
    """New rows of the form and a flag per row: new or changed."""
    rows: list[dict[str, str]] = []
    flags: list[bool] = []
    for index, row in enumerate(ordered, start=1):
        cells = {FORM: form, FILENAME: filename, COLUMN: column_letters(index), HEADER: row.name,
                 SOURCE: row.source, FIXED: row.fixed}
        if row.rid in old:
            record = {c: str(v) for c, v in table.iloc[row.rid].items()}
            changed = {h: v for h, v in cells.items() if _s(record[h]) != v}
            record.update(changed)
            if changed:
                record.update({UPDATED_AT: now, UPDATED_BY: user})
        else:
            record = {**dict.fromkeys(table.columns, ""), **cells, CHANNEL: LAYOUT_CHANNEL, UPDATED_AT: now, UPDATED_BY: user}
            changed = cells
        rows.append(record)
        flags.append(bool(changed))
    return rows, flags


def apply_edit(
    table: pd.DataFrame, text: str, form: str, filename: str, view0: pd.DataFrame, grid: pd.DataFrame,
    user: str, now: str,
) -> EditResult:
    """Fold the editor's result for one form onto the table and build the new CSV text."""
    positions = form_positions(table, form)
    filename = filename.strip()
    rows = grid_rows(grid)
    if _records(rows) == _records(grid_rows(view0)) and filename == form_filename(table, form):
        return EditResult(table, text, 0, 0, 0, tuple(positions), ())
    old = set(positions)
    ordered = final_order(rows)
    block, flags = _block(table, old, ordered, form, filename, user, now)
    survivors = {r.rid for r in ordered if r.rid in old}
    out: list[dict[str, str]] = []
    form_rows: list[int] = []
    for position, record in enumerate(table.to_dict("records")):
        if position not in old:
            out.append(record)
        elif position == positions[0]:
            form_rows += range(len(out), len(out) + len(block))
            out.extend(block)
    if not positions:
        form_rows = list(range(len(out), len(out) + len(block)))
        out.extend(block)
    new_table = pd.DataFrame(out, columns=list(table.columns))
    start = form_rows[0] if form_rows else 0
    changed = tuple(start + i for i, flag in enumerate(flags) if flag)
    added = sum(1 for r in ordered if r.rid not in old)
    return EditResult(
        new_table, to_csv_text(new_table), modified=sum(flags) - added, added=added,
        deleted=len(old) - len(survivors), form_positions=tuple(form_rows), changed_positions=changed,
    )


def delete_form(table: pd.DataFrame, text: str, form: str, user: str, now: str) -> EditResult:
    """The edit that removes every column of a form (the validator decides whether that is allowed)."""
    view0 = build_view(table, form)
    return apply_edit(table, text, form, form_filename(table, form), view0, view0.iloc[0:0], user, now)


def check_new_form(name: str, filename: str, existing: Sequence[str]) -> str | None:
    """Korean reason why this new form cannot be created, or None."""
    name, filename = name.strip(), filename.strip()
    if not name:
        return "양식 이름을 입력하세요."
    if name in existing:
        return f"'{name}' 양식이 이미 있습니다."
    bad = sorted({ch for ch in filename if ch in FORBIDDEN_FILENAME_CHARS})
    if bad:
        return f"파일명에 쓸 수 없는 글자가 있습니다: {' '.join(bad)}"
    return None


# ---------------------------------------------------------------- judging

@dataclass(frozen=True)
class Judgement:
    errors: list[str]
    warnings: list[str]
    reference: list[str]  # problems outside the edited rows (never block)
    edit: EditResult

    @property
    def has_changes(self) -> bool:
        return self.edit.has_changes


def _unique(lines: list[str]) -> list[str]:
    return list(dict.fromkeys(lines))


def judge(
    table: pd.DataFrame, text: str, form: str, filename: str, view0: pd.DataFrame, grid: pd.DataFrame,
    refs: LayoutReferences, route_csv: str, options_csv: str, user: str, now: str,
) -> Judgement:
    """Apply the edit and validate the result. Errors of the edited form block, warnings of new or changed
    rows need a confirmation, everything else is reference only. The final guard always applies."""
    edit = apply_edit(table, text, form, filename, view0, grid, user, now)
    is_new = form not in form_names(table)
    if not edit.has_changes and not is_new:
        return Judgement([], [], [], edit)
    issues = validate_layout(edit.table, table, refs, KNOWN_SOURCES, empty_forms=[form] if is_new else ())
    if edit.has_changes:
        issues += layout_final_guard(route_csv, options_csv, edit.text)
    errors: list[str] = []
    warnings: list[str] = []
    reference: list[str] = []
    for issue in issues:
        mine = issue.row is None or issue.row in edit.form_positions
        if issue.level == "error" and mine:
            errors.append(issue.message)
        elif issue.level == "warning" and (issue.row is None or issue.row in edit.changed_positions):
            warnings.append(issue.message)
        else:
            reference.append(issue.message)
    return Judgement(_unique(errors), _unique(warnings), _unique(reference), edit)


def summary_message(form: str, edit: EditResult) -> str:
    return f"발주서 양식 수정: {form} (수정 {edit.modified}, 추가 {edit.added}, 삭제 {edit.deleted})"
