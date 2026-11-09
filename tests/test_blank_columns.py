"""'(빈 칸)' spacer columns: empty header and cells, later columns keep the supplier form's place."""
from io import BytesIO

import openpyxl
import pandas as pd

from core.exporter import BLANK_HEADER, build_output_dataframe, export_individual_files
from store.layout_repo import COLUMN, FILENAME, FIXED, FORM, HEADER, SOURCE
from store.layout_validators import LayoutReferences, validate_layout

VENDOR = "업체"


def _layout(*columns: tuple[str, str, str, str]) -> pd.DataFrame:
    return pd.DataFrame(
        [{"VendorID": VENDOR, "FilePrefix": VENDOR, "ExcelCol": c, "HeaderName": h, "SourceCol": s,
          "HardcodedValue": f} for c, h, s, f in columns])


def _orders() -> pd.DataFrame:
    return pd.DataFrame({"_VendorID": [VENDOR, VENDOR], "수취인명": ["김가나", "이다라"]})


LAYOUT = _layout(("A", "보내는분", "", "최고다"), ("B", BLANK_HEADER, "", ""), ("C", BLANK_HEADER, "", ""),
                 ("D", "받는분", "수취인명", ""))


def test_blank_columns_are_empty_and_keep_positions():
    out = build_output_dataframe(_orders(), LAYOUT, VENDOR)
    assert list(out.columns) == ["보내는분", "", "", "받는분"]
    assert out.iloc[:, 3].tolist() == ["김가나", "이다라"]
    assert out.iloc[:, 1].tolist() == ["", ""]


def test_exported_sheet_has_blank_cells_in_place():
    file = export_individual_files(_orders(), {"OutputLayout": LAYOUT})[0]
    rows = list(openpyxl.load_workbook(BytesIO(file["data"].getvalue())).active.iter_rows(values_only=True))
    assert rows == [("보내는분", None, None, "받는분"), ("최고다", None, None, "김가나"), ("최고다", None, None, "이다라")]


def _validation(*rows: tuple[str, str, str, str]):
    frame = pd.DataFrame([{FORM: "양식", FILENAME: "양식", COLUMN: c, HEADER: h, SOURCE: s, FIXED: f,
                           "channel": "naver"} for c, h, s, f in rows])
    return validate_layout(frame, None, LayoutReferences())


def test_blank_header_may_repeat_but_other_names_may_not():
    assert _validation(("A", BLANK_HEADER, "", ""), ("B", BLANK_HEADER, "", ""), ("C", "받는분", "수취인명", "")) == []
    errors = _validation(("A", "받는분", "", ""), ("B", "받는분", "", ""))
    assert [i.level for i in errors] == ["error", "error"]


def test_content_in_a_blank_column_is_warned():
    issues = _validation(("A", BLANK_HEADER, "수취인명", ""))
    assert [i.level for i in issues] == ["warning"] and BLANK_HEADER in issues[0].message
