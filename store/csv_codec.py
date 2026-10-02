"""CSV <-> DataFrame conversion (UTF-8 with BOM, Excel friendly)."""
from __future__ import annotations

from io import StringIO

import pandas as pd

BOM = "﻿"


def to_csv_text(df: pd.DataFrame) -> str:
    """Serialize to CSV text starting with a BOM, no index, LF line endings."""
    return BOM + df.to_csv(index=False, lineterminator="\n")


def from_csv_text(text: str, as_text: bool = False) -> pd.DataFrame:
    """Parse CSV text with default type inference (mirrors pd.read_excel).

    as_text=True keeps every cell as the original string (no NA conversion, no number parsing).
    """
    body = StringIO(text.removeprefix(BOM))
    if as_text:
        return pd.read_csv(body, dtype=str, keep_default_na=False)
    return pd.read_csv(body)
