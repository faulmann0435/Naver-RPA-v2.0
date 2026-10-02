"""CSV codec: BOM, round trip, type inference like pd.read_excel."""
import numpy as np
import pandas as pd
import pandas.testing as pdt

from store.csv_codec import from_csv_text, to_csv_text


def test_bom_and_line_endings():
    text = to_csv_text(pd.DataFrame({"가": [1], "나": ["x"]}))
    assert text.startswith("﻿")
    assert "\r" not in text
    assert text.endswith("\n")


def test_round_trip_korean():
    df = pd.DataFrame({"우선순위": [0, 1], "키워드": ["코다리, 홍게", "a\nb"], "양식": ["속초", "문어"]})
    pdt.assert_frame_equal(from_csv_text(to_csv_text(df)), df)


def test_ints_inferred():
    out = from_csv_text("﻿a,b\n1,2\n3,4\n")
    assert out["a"].dtype.kind == "i"
    assert out.columns.tolist() == ["a", "b"]


def test_empty_becomes_nan():
    out = from_csv_text("a,b\n1,\n2,x\n")
    assert np.isnan(out.loc[0, "b"])
    assert out.loc[1, "b"] == "x"
