"""Option rule engine: apply_option_rules, run_option_engine."""
import re

import pandas as pd

from core.actions import (
    _apply_append_qty_unit,
    _apply_append_suffix,
    _apply_calc_unit,
    _apply_convert_weight_range_fix,
    _apply_format_qty_single_stamp,
    _apply_group_multiply,
    _apply_mask_text,
    _apply_prepend_text,
    _apply_remove_regex,
    _apply_remove_text,
    _apply_replace_regex_sub,
    _apply_unmask_text,
    _safe_int,
)


def apply_option_rules(row, option_rules, name_col="상품명", option_col="옵션정보", qty_col="수량", row_index=None, debug_log=None):
    current_vendor = str(row.get("_VendorID", "") or "").strip().upper()
    product = str(row.get(name_col, "") or "").strip()
    raw_option = row.get(option_col, "")
    if pd.isna(raw_option):
        raw_option = ""
    # text = option only (modify target); full_context = product + option (for keyword search only)
    text = str(raw_option).strip()
    full_context = f"{product} {text}".strip()
    qty = _safe_int(row.get(qty_col, 1), 1)
    calculated_weight = 0.0
    weight_calculated = False  # One-Shot Lock: prevents double CONVERT_WEIGHT application
    qty_display_lock = False   # v14.4: prevents double quantity suffix (x2 x2)
    calculated_weight_ref = [calculated_weight]
    weight_calculated_ref = [weight_calculated]
    qty_display_lock_ref = [qty_display_lock]
    do_log = debug_log is not None and row_index is not None and row_index < 5

    for rule_idx, (_, rule) in enumerate(option_rules.iterrows(), start=1):
        rule_vendor = rule.get("ApplyTo", "") or ""
        rule_target = rule.get("TargetKeyword", "") or ""
        action = str(rule.get("ActionType", "") or "").strip().upper()
        param = rule.get("Parameter", "") or ""

        if rule_vendor != "ALL" and rule_vendor != current_vendor:
            if do_log:
                debug_log.append(f"Row {row_index} Rule #{rule_idx} (Action: {action}, Param: {repr(param)[:50]}) -> Matched? NO (ApplyTo)")
            continue
        if rule_target != "ALL" and rule_target not in full_context:
            if do_log:
                debug_log.append(f"Row {row_index} Rule #{rule_idx} (Action: {action}, Param: {repr(param)[:50]}) -> Matched? NO (TargetKeyword)")
            continue

        # CONVERT_WEIGHT One-Shot Lock: skip if already calculated for this row
        if action == "CONVERT_WEIGHT":
            if weight_calculated_ref[0]:
                if do_log:
                    debug_log.append(f"Row {row_index} Rule #{rule_idx} (Action: CONVERT_WEIGHT) -> SKIP (lock)")
                continue
            text = _apply_convert_weight_range_fix(text, qty, calculated_weight_ref)
            weight_calculated_ref[0] = True
        elif action == "REMOVE_TEXT":
            text = _apply_remove_text(text, param)
        elif action == "REMOVE_REGEX":
            text = _apply_remove_regex(text, param)
        elif action == "REPLACE_REGEX_SUB":
            text = _apply_replace_regex_sub(text, param)
        elif action == "MASK_TEXT":
            text = _apply_mask_text(text, param)
        elif action == "UNMASK_TEXT":
            text = _apply_unmask_text(text, param)
        elif action == "CALC_UNIT":
            prev_text = text
            text = _apply_calc_unit(text, param, qty)
            if text != prev_text and qty > 1:
                qty_display_lock_ref[0] = True
        elif action == "APPEND_QTY_UNIT":
            if qty_display_lock_ref[0]:
                if do_log:
                    debug_log.append(f"Row {row_index} Rule #{rule_idx} (Action: APPEND_QTY_UNIT) -> SKIP (qty_display_lock)")
                continue
            text = _apply_append_qty_unit(text, param, qty)
            qty_display_lock_ref[0] = True
        elif action == "GROUP_MULTIPLY":
            if qty_display_lock_ref[0]:
                if do_log:
                    debug_log.append(f"Row {row_index} Rule #{rule_idx} (Action: GROUP_MULTIPLY) -> SKIP (qty_display_lock)")
                continue
            # Trigger if param in full_context; append (xQty) to text (option only)
            prev_text = text
            text = _apply_group_multiply(text, param, qty, search_in=full_context)
            if text != prev_text:
                qty_display_lock_ref[0] = True
        elif action == "APPEND_SUFFIX":
            text = _apply_append_suffix(text, param)
        elif action == "PREPEND_TEXT":
            text = _apply_prepend_text(text, param)
        elif action == "FORMAT_QTY":
            if qty_display_lock_ref[0]:
                if do_log:
                    debug_log.append(f"Row {row_index} Rule #{rule_idx} (Action: FORMAT_QTY) -> SKIP (qty_display_lock)")
                continue
            text = _apply_format_qty_single_stamp(text, param, qty, qty_display_lock_ref)

        text = re.sub(r"\s+", " ", str(text)).strip()
        if do_log:
            debug_log.append(f"Row {row_index} Rule #{rule_idx} (Action: {action}, Param: {repr(param)[:50]}) -> Matched? YES -> Result: {repr(text)[:50]}")

    final_weight = calculated_weight_ref[0]
    final_formatted = qty_display_lock_ref[0]  # True if any qty stamp was applied
    if not text:
        text = f" x{qty}개" if not final_formatted else " "
        text = text.strip() or f"{qty}개"
    return (text.strip(), final_weight, final_formatted)


def run_option_engine(df, option_rules, debug_log=None):
    # Column detection: Korean first, fallback to English (e.g. renamed headers)
    name_col = "상품명" if "상품명" in df.columns else ("Product Name" if "Product Name" in df.columns else None)
    option_col = "옵션정보" if "옵션정보" in df.columns else ("Option" if "Option" in df.columns else None)
    qty_col = "수량" if "수량" in df.columns else ("Qty" if "Qty" in df.columns else None)
    if not name_col:
        raise ValueError("상품명 또는 'Product Name' 컬럼이 없습니다. 데이터에 제품명 컬럼을 추가하세요.")
    if not option_col:
        raise ValueError("옵션정보 또는 'Option' 컬럼이 없습니다. 데이터에 옵션 컬럼을 추가하세요.")
    if not qty_col:
        raise ValueError("수량 또는 'Qty' 컬럼이 없습니다. 데이터에 수량 컬럼을 추가하세요.")
    df = df.copy()
    df["_calculated_weight"] = 0.0
    df["_is_formatted"] = False
    df["_weight_calculated"] = False  # One-Shot Lock (per-row, used in apply_option_rules)
    # List-collection pattern (v14.1): collect in lists then assign once (faster than df.at[i] per row)
    opts = []
    weights = []
    formatted_flags = []
    for i in range(len(df)):
        row = df.iloc[i].copy()
        r_text, r_weight, r_fmt = apply_option_rules(
            row, option_rules, name_col, option_col, qty_col, row_index=i, debug_log=debug_log
        )
        opts.append(r_text)
        weights.append(r_weight)
        formatted_flags.append(r_fmt)
    df["processed_option"] = opts
    df["_calculated_weight"] = weights
    df["_is_formatted"] = formatted_flags
    return df
