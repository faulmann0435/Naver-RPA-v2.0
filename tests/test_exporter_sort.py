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


def test_row_number_source_fills_1_to_n():
    from core.exporter import ROW_NUMBER_SOURCE

    layout = pd.DataFrame({
        "VendorID": ["v", "v", "v"], "ExcelCol": ["A", "B", "C"], "HeaderName": ["순번", "이름", "고정"],
        "SourceCol": [ROW_NUMBER_SOURCE, "name", ROW_NUMBER_SOURCE], "HardcodedValue": ["", "", "x"],
    })
    out = build_output_dataframe(pd.DataFrame({"name": ["가", "나", "다"]}, index=[7, 3, 9]), layout, "v")
    assert out["순번"].tolist() == [1, 2, 3] and out["이름"].tolist() == ["가", "나", "다"]
    assert out["고정"].tolist() == ["x", "x", "x"]  # fixed text still wins
