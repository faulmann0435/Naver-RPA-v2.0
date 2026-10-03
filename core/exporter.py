"""Export: output dataframe building, Jeju highlight, per-vendor Excel files."""
from datetime import datetime
from io import BytesIO

import pandas as pd

try:
    from xlsxwriter.utility import xl_col_to_name
except ImportError:
    xl_col_to_name = None


# 매핑데이터 value that fills a column with 1, 2, 3 ... per output file (not a column of the order file)
ROW_NUMBER_SOURCE = "_순번"


def excel_column_number(value) -> int | None:
    """Excel column letters as a number (A=1 ... Z=26, AA=27); None when it is not a letter code."""
    text = str(value).strip().upper() if pd.notna(value) else ""
    if not text or not (text.isascii() and text.isalpha()):
        return None
    number = 0
    for char in text:
        number = number * 26 + (ord(char) - ord("A") + 1)
    return number


def _sort_by_excel_column(layout):
    """Stable order by column letter value (so AA comes after Z); non-letter values last."""
    keys = [excel_column_number(v) for v in layout["ExcelCol"]]
    order = sorted(range(len(keys)), key=lambda i: (keys[i] is None, keys[i] or 0))
    return layout.iloc[order].reset_index(drop=True)


def build_output_dataframe(merged_df, output_layout, vendor_id):
    layout = output_layout[output_layout["VendorID"].astype(str).str.strip() == str(vendor_id).strip()]
    if layout.empty:
        return None
    layout = _sort_by_excel_column(layout)
    row_count = len(merged_df)
    out = {}
    for _, r in layout.iterrows():
        header = str(r["HeaderName"]).strip() if pd.notna(r["HeaderName"]) else ""
        if not header:
            continue
        hc = r.get("HardcodedValue")
        if pd.notna(hc) and str(hc).strip():
            out[header] = [str(hc).strip()] * row_count
        else:
            src = r.get("SourceCol")
            if pd.notna(src) and str(src).strip() == ROW_NUMBER_SOURCE:
                out[header] = list(range(1, row_count + 1))
            elif pd.notna(src) and str(src).strip() and str(src).strip() in merged_df.columns:
                out[header] = merged_df[str(src).strip()].values
            else:
                out[header] = [""] * row_count
    if not out:
        return None
    return pd.DataFrame(out)


# Keywords for address-related columns (for conditional formatting "제주" highlight)
_ADDRESS_COLUMN_KEYWORDS = ("주소", "배송지", "address")


def _apply_jeju_highlight(writer, out_df, sheet_name="발주"):
    """Apply conditional format: light/dark red for cells containing '제주' in address-related columns only."""
    workbook = writer.book
    worksheet = writer.sheets.get(sheet_name)
    if worksheet is None:
        return
    jeju_fmt = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
    nrows = len(out_df) + 1  # +1 for header; data rows 2..nrows
    for col_idx, col_name in enumerate(out_df.columns):
        name_str = str(col_name).strip()
        if not name_str:
            continue
        is_address_col = any(
            (kw in name_str) if kw != "address" else (kw in name_str.lower())
            for kw in _ADDRESS_COLUMN_KEYWORDS
        )
        if not is_address_col:
            continue
        col_letter = xl_col_to_name(col_idx)
        cell_range = f"{col_letter}2:{col_letter}{nrows}"
        worksheet.conditional_format(cell_range, {
            "type": "text",
            "criteria": "containing",
            "value": "제주",
            "format": jeju_fmt,
        })


def export_individual_files(merged_df, config):
    """Generate one Excel file per vendor. Returns list of {vendor, data, filename} (no ZIP)."""
    output_layout = config["OutputLayout"]
    vendor_ids = merged_df["_VendorID"].dropna().astype(str).str.strip().unique()
    vendor_ids = [v for v in vendor_ids if v and v != "Unclassified"]

    date_str = datetime.now().strftime("%Y%m%d")
    processed_files = []

    for vid in vendor_ids:
        subset = merged_df[merged_df["_VendorID"].astype(str).str.strip() == vid]
        layout = config["OutputLayout"]
        prefix_row = layout[layout["VendorID"].astype(str).str.strip() == vid]
        file_prefix = ""
        if not prefix_row.empty and "FilePrefix" in prefix_row.columns:
            file_prefix = str(prefix_row["FilePrefix"].iloc[0]).strip() if pd.notna(prefix_row["FilePrefix"].iloc[0]) else ""
        if file_prefix and str(file_prefix).lower().endswith(".xlsx"):
            file_prefix = str(file_prefix)[:-5]
        file_prefix = str(file_prefix).strip() if file_prefix else ""
        filename = f"{file_prefix}_{date_str}.xlsx" if file_prefix else f"{vid}_{date_str}.xlsx"

        out_df = build_output_dataframe(subset, layout, vid)
        if out_df is None or out_df.empty:
            out_df = subset.copy()
        excel_buf = BytesIO()
        with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as writer:
            out_df.to_excel(writer, index=False, sheet_name="발주")
            # Conditional format: highlight "제주" only in address-related columns
            if xl_col_to_name and not out_df.empty:
                _apply_jeju_highlight(writer, out_df, sheet_name="발주")
        excel_buf.seek(0)

        processed_files.append({
            "vendor": vid,
            "data": excel_buf,
            "filename": filename,
        })

    return processed_files
