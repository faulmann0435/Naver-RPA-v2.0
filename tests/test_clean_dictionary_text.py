import pandas as pd

from store.rules_repo import DICTIONARY_COLUMNS
from tools.clean_dictionary_text import cleaned_frame


def test_cleaned_frame_only_touches_text_fields_and_stamps_rows():
    rows = [
        {"product_name_ref": "🔥홍게", "option_raw_ref": "중량: 🔥라면용 홍게 3kg", "display_template": "🔥라면용 홍게 3kg",
         "sum_group": "", "display_template_qty1": ""},
        {"product_name_ref": "문어", "option_raw_ref": "x", "display_template": "", "sum_group": "★제수용 피문어",
         "display_template_qty1": ""},
    ]
    frame = pd.DataFrame([{**dict.fromkeys(DICTIONARY_COLUMNS, ""), **r} for r in rows], columns=DICTIONARY_COLUMNS)
    new, cells, changed = cleaned_frame(frame, "clean tool", "2026-10-02T00:00:00+09:00")
    assert (cells, changed) == (1, 1)
    assert new.at[0, "display_template"] == "라면용 홍게 3kg"
    assert new.at[0, "product_name_ref"] == "🔥홍게" and new.at[0, "option_raw_ref"].startswith("중량: 🔥")
    assert new.at[0, "updated_by"] == "clean tool" and new.at[1, "updated_by"] == ""
    assert new.at[1, "sum_group"] == "★제수용 피문어"
