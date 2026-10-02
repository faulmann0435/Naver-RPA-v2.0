"""Merge & sort: instruction-row filter, order merge, payment-date sort."""
import re

import pandas as pd

from core.actions import _safe_int

FILTER_PHRASE = "다운로드 받은 파일로 '엑셀 일괄발송' 처리하는 방법"


def filter_instruction_rows(df):
    if df.empty:
        return df
    mask = df.astype(str).apply(
        lambda row: row.str.contains(FILTER_PHRASE, na=False).any(), axis=1
    )
    return df.loc[~mask].reset_index(drop=True)


def _cleanup_empty_parens(s):
    """Remove empty parentheses () or [] that might remain after removals."""
    if pd.isna(s) or not str(s).strip():
        return s
    t = str(s)
    t = re.sub(r"\(\s*\)", "", t)
    t = re.sub(r"\[\s*\]", "", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def merge_orders(df, option_rules=None):
    bundle_col = "배송비 묶음번호"
    if bundle_col not in df.columns:
        raise ValueError(f"병합 키 컬럼 없음: {bundle_col}")
    # 배송비 묶음번호 기준 병합 + _VendorID(내부 라우팅용)
    group_cols = [bundle_col, "_VendorID"]
    for c in group_cols:
        if c not in df.columns:
            raise ValueError(f"병합 키 컬럼 없음: {c}")

    def join_unique_messages(ser):
        parts = ser.dropna().astype(str).str.strip().unique().tolist()
        return " / ".join(str(p) for p in parts if p)

    def _format_weight(total_weight):
        if total_weight < 1.0:
            return f"{round(total_weight * 1000)}g"
        rounded = round(total_weight, 3)
        if rounded == int(rounded):
            return f"{int(rounded)}kg"
        return f"{rounded}kg"

    def process_group(gdf):
        has_weight_col = "_calculated_weight" in gdf.columns
        has_qty_col = "수량" in gdf.columns
        weight_vals = gdf["_calculated_weight"].fillna(0) if has_weight_col else pd.Series(0.0, index=gdf.index)

        weight_dict = {}   # option -> total_weight (insertion order preserved)
        normal_dict = {}   # option -> total_qty (insertion order preserved)

        for idx, row in gdf.iterrows():
            w = float(weight_vals.loc[idx]) if has_weight_col else 0.0
            opt = row.get("processed_option")
            processed_option = "" if (opt is None or (isinstance(opt, float) and pd.isna(opt))) else str(opt).strip()
            is_formatted = bool(row.get("_is_formatted", False))
            row_qty = (1 if is_formatted else _safe_int(row.get("수량"), 1)) if has_qty_col else 1
            if w > 0:
                weight_dict[processed_option] = weight_dict.get(processed_option, 0) + w
            else:
                if processed_option:
                    normal_dict[processed_option] = normal_dict.get(processed_option, 0) + row_qty

        # 무게 옵션: 동일 옵션은 무게 합산 (기존 로직 유지)
        formatted_weight_strings = [f"{name} {_format_weight(t)}" for name, t in weight_dict.items()]
        weight_str = " / ".join(str(x) for x in formatted_weight_strings) if formatted_weight_strings else ""
        # 일반 옵션: 동일 옵션은 "/" 없이 단일 표시 + (x qty) 표기
        normal_parts = []
        for opt, total_qty in normal_dict.items():
            if total_qty > 1:
                normal_parts.append(f"{opt} (x{total_qty})")
            else:
                normal_parts.append(opt)
        normal_str = " / ".join(str(x) for x in normal_parts) if normal_parts else ""
        if weight_str and normal_str:
            processed_option = weight_str + " / " + normal_str
        else:
            processed_option = weight_str or normal_str

        total_weight = sum(weight_dict.values()) if has_weight_col else 0.0

        out = {}
        for col in gdf.columns:
            if col in group_cols:
                out[col] = gdf[col].iloc[0]
            elif col == "processed_option":
                out[col] = processed_option
            elif col == "_calculated_weight":
                out[col] = total_weight if has_weight_col else 0.0
            elif col == "수량":
                out[col] = gdf[col].apply(lambda v: _safe_int(v, 1)).sum()
            elif col == "배송메세지":
                out[col] = join_unique_messages(gdf[col])
            elif col == "결제일":
                out[col] = gdf[col].min()
            elif col == "구매자명":
                out[col] = gdf[col].iloc[0]
            else:
                out[col] = gdf[col].iloc[0]

        return pd.Series(out)

    merged = df.groupby(group_cols, as_index=False).apply(process_group, include_groups=False)
    merged["processed_option"] = merged["processed_option"].apply(_cleanup_empty_parens)
    return merged


def sort_by_payment_date(df):
    if "결제일" not in df.columns or df.empty:
        return df
    df = df.copy()
    s = pd.to_datetime(df["결제일"], errors="coerce")
    df["_sort_date"] = s
    df = df.sort_values("_sort_date", ascending=True, na_position="last").drop(columns=["_sort_date"])
    return df.reset_index(drop=True)
