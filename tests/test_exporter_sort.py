"""build_output_dataframe orders columns by the numeric value of the column letters."""
import pandas as pd

from core.exporter import build_output_dataframe, excel_column_number


def _layout(cols: list[str]) -> pd.DataFrame:
    return pd.DataFrame({
        "VendorID": ["v"] * len(cols), "ExcelCol": cols, "HeaderName": [f"h{c}" for c in cols],
        "SourceCol": [""] * len(cols), "HardcodedValue": ["x"] * len(cols),
    })


def test_column_number():
    assert [excel_column_number(v) for v in ("A", "Z", "AA", "ab", " b ")] == [1, 26, 27, 28, 2]
    assert excel_column_number("") is None and excel_column_number("1") is None
    assert excel_column_number(float("nan")) is None


def test_columns_past_z_come_after_z():
    out = build_output_dataframe(pd.DataFrame({"a": [1]}), _layout(["AA", "B", "Z", "A"]), "v")
    assert list(out.columns) == ["hA", "hB", "hZ", "hAA"]


def test_non_letter_values_last_and_stable():
    out = build_output_dataframe(pd.DataFrame({"a": [1]}), _layout(["?", "B", "", "A", "!"]), "v")
    assert list(out.columns) == ["hA", "hB", "h?", "h", "h!"]
