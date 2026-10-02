"""CSV <-> DataFrame conversion (UTF-8 with BOM, Excel friendly)."""
from __future__ import annotations

from io import StringIO

import pandas as pd

BOM = "﻿"


def to_csv_text(df: pd.DataFrame) -> str:
    """Serialize to CSV text starting with a BOM, no index, LF line endings."""
    return BOM + df.to_csv(index=False, lineterminator="\n")


def from_csv_text(text: str) -> pd.DataFrame:
    """Parse CSV text with default type inference (mirrors pd.read_excel)."""
    return pd.read_csv(StringIO(text.removeprefix(BOM)))
