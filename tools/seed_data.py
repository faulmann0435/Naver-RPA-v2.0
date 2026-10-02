"""Data loading for the seed tool: order files, past purchase-order (PO) files and their join.

Order and PO files hold customer data; only product-level fields leave this module.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pandas as pd

from core.dictionary import DictionarySettings
from core.loader import (
    _get_excel_bytes,
    _strip_columns,
    ensure_quantity_column,
    find_header_row,
)
from core.option_key import make_option_key, normalize_product_no, normalize_text
from core.seed import HandEvidence, safe_qty, strip_ignored_groups

ORDER_SHEET = "발주발송관리"
REQUIRED_ORDER_COLUMNS = frozenset({"상품명", "옵션정보", "수량", "배송비 묶음번호", "상품번호"})
ORDER_ID_COLUMN = "상품주문번호"
PO_NAME_COL = "행 레이블"
PO_PHONE_COL = "수취인연락처1"
PO_OPTION_COL = "옵션정보"
PO_HAND_POSITION = 5
ORDER_NAME_COL = "수취인명"
_PO_NAME = re.compile(r"^(?P<prefix>.+?) 발주서 (?P<rest>.+)\.xlsx$")
_NON_DIGITS = re.compile(r"\D")


class _Bytes:
    """Minimal upload-like object for core.loader._get_excel_bytes."""

    def __init__(self, path: Path) -> None:
        self._data = path.read_bytes()

    def read(self) -> bytes:
        return self._data


def read_xlsx_bytes(path: Path, password: str | None) -> bytes | None:
    """Plain bytes of an xlsx (decrypted when needed); None when unreadable or encrypted without a password."""
    try:
        return _get_excel_bytes(_Bytes(path), None)
    except ValueError:
        if not password:
            return None
    try:
        return _get_excel_bytes(_Bytes(path), password)
    except ValueError:
        return None


def _read_frame(raw: bytes, sheet: str | int) -> pd.DataFrame:
    preview = pd.read_excel(BytesIO(raw), sheet_name=sheet, header=None, nrows=20)
    header_idx = find_header_row(preview)
    frame = pd.read_excel(BytesIO(raw), sheet_name=sheet, header=header_idx)
    _strip_columns(frame)
    ensure_quantity_column(frame)
    return frame


def load_order_frame(path: Path, password: str | None) -> pd.DataFrame | None:
    """Order table of one file (sheet 발주발송관리 when present, else the first sheet); None if not an order file."""
    raw = read_xlsx_bytes(path, password)
    if raw is None:
        return None
    try:
        sheets = pd.ExcelFile(BytesIO(raw)).sheet_names
        frame = _read_frame(raw, ORDER_SHEET if ORDER_SHEET in sheets else 0)
    except Exception:  # noqa: BLE001 - any unreadable workbook just means 'not an order file'
        return None
    return frame if REQUIRED_ORDER_COLUMNS <= set(frame.columns) else None


def discover_order_files(root: Path, repo_root: Path, exclude: tuple[Path, ...] = ()) -> list[Path]:
    found = []
    for path in sorted(root.rglob("*.xlsx")):
        if repo_root in path.parents or path.name.startswith("~$") or path.name == "config.xlsx":
            continue
        if any(skipped in path.parents for skipped in exclude):
            continue
        found.append(path)
    return found


@dataclass(frozen=True)
class OrdersLoad:
    orders: pd.DataFrame
    files_read: int
    files_skipped: int
    rows_total: int
    rows_duplicate: int


def load_all_orders(paths: list[Path], password: str | None) -> OrdersLoad:
    """Every readable order file, de-duplicated by 상품주문번호 (first kept)."""
    frames = []
    skipped = 0
    for path in paths:
        frame = load_order_frame(path, password)
        if frame is None:
            skipped += 1
        else:
            frames.append(frame)
    if not frames:
        return OrdersLoad(pd.DataFrame(columns=sorted(REQUIRED_ORDER_COLUMNS)), 0, skipped, 0, 0)
    orders = pd.concat(frames, ignore_index=True)
    total = len(orders)
    if ORDER_ID_COLUMN in orders.columns:
        ids = orders[ORDER_ID_COLUMN]
        duplicate = ids.notna() & ids.astype(str).duplicated(keep="first")
        orders = orders[~duplicate].reset_index(drop=True)
    return OrdersLoad(orders, len(frames), skipped, total, total - len(orders))


@dataclass(frozen=True)
class PoPair:
    po_path: Path
    order_path: Path

    @property
    def label(self) -> str:
        return self.po_path.stem


def find_po_pairs(folder: Path) -> list[PoPair]:
    """`<prefix> 발주서 <rest>.xlsx` pairs with `<prefix> 주문서 <rest>.xlsx` in the same folder."""
    pairs = []
    for po_path in sorted(folder.glob("*.xlsx")):
        match = _PO_NAME.match(po_path.name)
        if not match:
            continue
        order_path = folder / f"{match['prefix']} 주문서 {match['rest']}.xlsx"
        if order_path.exists():
            pairs.append(PoPair(po_path, order_path))
    return pairs


def load_po_frame(path: Path, password: str | None) -> pd.DataFrame | None:
    """PO rows with columns name, phone, option, hand (hand text = 6th column); None when unreadable."""
    raw = read_xlsx_bytes(path, password)
    if raw is None:
        return None
    try:
        book = pd.ExcelFile(BytesIO(raw))
        for sheet in book.sheet_names:
            frame = book.parse(sheet, header=0)
            _strip_columns(frame)
            if PO_NAME_COL in frame.columns and frame.shape[1] > PO_HAND_POSITION:
                return _po_columns(frame)
    except Exception:  # noqa: BLE001 - any unreadable workbook just means 'no PO table'
        return None
    return None


def clean_text(value: object) -> str:
    """Stripped text of a cell; empty for None / NaN / NaT / NA."""
    if value is None or value is pd.NA or value is pd.NaT:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def _po_columns(frame: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame({
        "name": frame[PO_NAME_COL].map(clean_text),
        "phone": frame[PO_PHONE_COL].map(clean_text),
        "option": frame[PO_OPTION_COL].map(clean_text),
        "hand": frame.iloc[:, PO_HAND_POSITION].map(clean_text),
    })
    return out[(out["name"] != "") & (out["name"] != "총합계")].reset_index(drop=True)


@dataclass(frozen=True)
class PoBundle:
    """One recipient of a PO: `index` is its position in the file (1-based); the name stays internal."""

    index: int
    name: str
    hand: str


@dataclass(frozen=True)
class PoJoin:
    bundles: list[PoBundle]
    evidence: dict[tuple[str, str], list[HandEvidence]]
    po_rows: int
    unmatched_rows: int
    ambiguous_rows: int
    no_hand_bundles: int


def _digits(value: object) -> str:
    return _NON_DIGITS.sub("", clean_text(value))


def _bundles(po: pd.DataFrame) -> tuple[list[PoBundle], int]:
    groups: dict[tuple[str, str], list[str]] = {}
    for name, phone, hand in zip(po["name"], po["phone"], po["hand"]):
        groups.setdefault((name, _digits(phone)), []).append(hand)
    bundles: list[PoBundle] = []
    no_hand = 0
    for index, ((name, _), hands) in enumerate(groups.items(), start=1):
        hand = next((h for h in hands if h), "")
        if hand:
            bundles.append(PoBundle(index, name, hand))
        else:
            no_hand += 1
    return bundles, no_hand


def _join_keys(frame: pd.DataFrame, name: str, phone: str, option: str) -> pd.DataFrame:
    return pd.DataFrame({
        "_name": frame[name].map(clean_text),
        "_phone": frame[phone].map(_digits),
        "_opt": frame[option].map(lambda v: normalize_text(clean_text(v))),
    })


def join_po_to_orders(po: pd.DataFrame, orders: pd.DataFrame, settings: DictionarySettings) -> PoJoin:
    """Join PO rows to order rows on (수취인명, 연락처1, 옵션정보) and collect single-row-bundle evidence."""
    bundles, no_hand = _bundles(po)
    required = {ORDER_NAME_COL, PO_PHONE_COL}
    if not required <= set(orders.columns):
        return PoJoin(bundles, {}, len(po), len(po), 0, no_hand)
    order_keys = _join_keys(orders, ORDER_NAME_COL, PO_PHONE_COL, "옵션정보")
    order_keys["_oid"] = range(len(orders))
    po_keys = _join_keys(po, "name", "phone", "option")
    po_keys["_pid"] = range(len(po))
    merged = po_keys.merge(order_keys, how="left", on=["_name", "_phone", "_opt"])
    per_row = merged.groupby("_pid")["_oid"].agg(lambda s: int(s.notna().sum()))
    unmatched = int((per_row == 0).sum())
    ambiguous = int((per_row > 1).sum())

    bundle_ids = orders["배송비 묶음번호"].map(clean_text)
    bundle_sizes = bundle_ids[bundle_ids != ""].value_counts()
    evidence: dict[tuple[str, str], list[HandEvidence]] = {}
    unique = merged[merged["_pid"].map(per_row) == 1].dropna(subset=["_oid"])
    for pid, oid in zip(unique["_pid"], unique["_oid"].astype(int)):
        hand = str(po["hand"].iat[int(pid)])
        order = orders.iloc[oid]
        bundle = bundle_ids.iat[oid]
        if not hand or not bundle or bundle_sizes.get(bundle, 0) != 1:
            continue
        option = clean_text(order["옵션정보"])
        key = (
            normalize_product_no(order["상품번호"]),
            make_option_key(option, settings.ignored_option_groups),
        )
        item = HandEvidence(hand, safe_qty(order["수량"]), strip_ignored_groups(option, settings.ignored_option_groups))
        evidence.setdefault(key, []).append(item)
    return PoJoin(bundles, evidence, len(po), unmatched, ambiguous, no_hand)
