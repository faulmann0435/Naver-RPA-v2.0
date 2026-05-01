"""
Sokcho Order Processing System v14.4
- Quantity Display Lock: prevents double quantity stamping (e.g., "x2 x2")
- GROUP_MULTIPLY: append (x{Qty}) only, no param word repeat
- Config: load_config_local, List-based, Session State
"""
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    import msoffcrypto
    HAS_MSOFFCRYPTO = True
except ImportError:
    HAS_MSOFFCRYPTO = False

try:
    from xlsxwriter.utility import xl_col_to_name
except ImportError:
    xl_col_to_name = None

# --- Column names (order data) ---
HEADER_KEYWORDS = ["상품명", "수취인명", "옵션정보"]
FILTER_PHRASE = "다운로드 받은 파일로 '엑셀 일괄발송' 처리하는 방법"
QTY_COLUMN_INDEX = 12
DEFAULT_PHONE_COL = "수취인연락처1"
ALT_PHONE_COL = "구매자연락처"


def _strip_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_header_row(preview_df):
    for i in range(len(preview_df)):
        row_vals = preview_df.iloc[i].astype(str).str.strip().tolist()
        row_text = " ".join(str(x) for x in row_vals)
        if all(kw in row_text for kw in HEADER_KEYWORDS):
            return i
    return 0


def ensure_quantity_column(df):
    if "수량" in df.columns:
        return df
    if df.shape[1] <= QTY_COLUMN_INDEX:
        return df
    df.rename(columns={df.columns[QTY_COLUMN_INDEX]: "수량"}, inplace=True)
    return df


# ============== Config Loader (v14.3: load_config_local + cache) ==============

@st.cache_data(ttl=3600)
def load_config_local(config_path: str, _password: str = None, cache_key: str = None):
    """
    Load config.xlsx with password support. Uses msoffcrypto if encrypted.
    Returns: dict with keys ['ProductRoute', 'OptionRules', 'OutputLayout']
    _cache_key: pass file mtime/size to invalidate cache when file changes.
    """
    path = Path(config_path).resolve()
    if not path.exists():
        raise FileNotFoundError(
            f"설정 파일을 찾을 수 없습니다. 경로: {path!s}\n"
            "config.xlsx를 앱과 같은 폴더에 두거나 경로를 확인하세요."
        )
    password = _password if _password and str(_password).strip() else None
    # _cache_key used only for cache invalidation (not for decryption)

    raw = path.read_bytes()
    if HAS_MSOFFCRYPTO:
        bio = BytesIO(raw)
        office_file = msoffcrypto.OfficeFile(bio)
        if office_file.is_encrypted():
            if not (password and str(password).strip()):
                raise ValueError("설정 파일이 암호화되어 있습니다. 비밀번호를 입력하세요.")
            try:
                office_file.load_key(password=str(password).strip())
                decrypted = BytesIO()
                office_file.decrypt(decrypted)
                decrypted.seek(0)
                raw = decrypted.getvalue()
            except Exception as e:
                raise ValueError("설정 파일 비밀번호가 올바르지 않습니다.") from e

    xl = pd.ExcelFile(BytesIO(raw))
    required_sheets = ["ProductRoute", "OptionRules", "OutputLayout"]
    missing = [s for s in required_sheets if s not in xl.sheet_names]
    if missing:
        raise ValueError(f"설정 시트 누락: {missing}. 필요: {required_sheets}")

    product_route = pd.read_excel(xl, sheet_name="ProductRoute")
    option_rules = pd.read_excel(xl, sheet_name="OptionRules")
    output_layout = pd.read_excel(xl, sheet_name="OutputLayout")

    _strip_columns(product_route)
    _strip_columns(option_rules)
    _strip_columns(output_layout)

    _debug_option_raw_headers = option_rules.columns.tolist()

    def _norm_product_route(df):
        rename_map = {}
        for col in df.columns:
            c = str(col).strip()
            if "우선순위" in c:
                rename_map[col] = "Priority"
            elif "키워드" in c:
                rename_map[col] = "Keyword"
            elif "양식명칭" in c:
                rename_map[col] = "TargetVendorID"
        return df.rename(columns=rename_map) if rename_map else df

    def _norm_option_rules(df):
        rename_map = {}
        for col in df.columns:
            c = str(col).strip()
            if "순서" in c:
                rename_map[col] = "Order"
            elif "설정값" in c or "Parameter" in c:
                rename_map[col] = "Parameter"
            elif "명령" in c or "Action" in c or "ActionType" in c:
                rename_map[col] = "ActionType"
            elif "적용대상" in c or "Target" in c or "TargetKeyword" in c:
                rename_map[col] = "TargetKeyword"
            elif "양식" in c or "Apply" in c or "양식명칭" in c:
                rename_map[col] = "ApplyTo"
            elif "Description" in c or "설명" in c:
                rename_map[col] = "Description"
        return df.rename(columns=rename_map) if rename_map else df

    def _norm_output_layout(df):
        rename_map = {}
        for col in df.columns:
            c = str(col).strip()
            if "양식명칭" in c:
                rename_map[col] = "VendorID"
            elif "파일명" in c:
                rename_map[col] = "FilePrefix"
            elif "열" in c:
                rename_map[col] = "ExcelCol"
            elif "헤더명" in c:
                rename_map[col] = "HeaderName"
            elif "매핑데이터" in c:
                rename_map[col] = "SourceCol"
            elif "고정값" in c:
                rename_map[col] = "HardcodedValue"
        return df.rename(columns=rename_map) if rename_map else df

    def _strip_all_strings(df):
        for col in df.columns:
            if df[col].dtype == object or df[col].dtype.kind == "O":
                df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        return df

    product_route = _norm_product_route(product_route)
    option_rules = _norm_option_rules(option_rules)
    output_layout = _norm_output_layout(output_layout)

    product_route = _strip_all_strings(product_route)
    option_rules = _strip_all_strings(option_rules)
    output_layout = _strip_all_strings(output_layout)

    # Data Cleaning: Parameter string, ApplyTo Uppercase
    if "Parameter" in option_rules.columns:
        option_rules["Parameter"] = option_rules["Parameter"].astype(str)
        option_rules["Parameter"] = option_rules["Parameter"].replace(["nan", "NaN", "None", "<NA>"], "")
        option_rules["Parameter"] = option_rules["Parameter"].str.strip()
    else:
        option_rules["Parameter"] = ""
    if "ApplyTo" in option_rules.columns:
        option_rules["ApplyTo"] = option_rules["ApplyTo"].fillna("").astype(str).str.strip().str.upper()
    else:
        option_rules["ApplyTo"] = ""
    if "TargetKeyword" in option_rules.columns:
        option_rules["TargetKeyword"] = option_rules["TargetKeyword"].fillna("").astype(str).str.strip()
    else:
        option_rules["TargetKeyword"] = ""

    _debug_option_renamed_headers = option_rules.columns.tolist()

    for name, df in [("ProductRoute", product_route), ("OptionRules", option_rules), ("OutputLayout", output_layout)]:
        if df.empty:
            raise ValueError(f"설정 시트 '{name}'이 비어 있습니다.")

    for col in ["Priority", "Keyword", "TargetVendorID"]:
        if col not in product_route.columns:
            raise ValueError(f"ProductRoute에 '{col}' 컬럼이 없습니다.")
    product_route = product_route.sort_values("Priority", ascending=True).reset_index(drop=True)

    for col in ["Order", "ApplyTo", "TargetKeyword", "ActionType"]:
        if col not in option_rules.columns:
            raise ValueError(f"OptionRules에 '{col}' 컬럼이 없습니다.")
    if "Parameter" not in option_rules.columns:
        option_rules["Parameter"] = ""
    if "Description" not in option_rules.columns:
        option_rules["Description"] = ""
    option_rules = option_rules.sort_values("Order", ascending=True).reset_index(drop=True)

    for col in ["VendorID", "FilePrefix", "ExcelCol", "HeaderName"]:
        if col not in output_layout.columns:
            raise ValueError(f"OutputLayout에 '{col}' 컬럼이 없습니다.")
    if "SourceCol" not in output_layout.columns:
        output_layout["SourceCol"] = ""
    if "HardcodedValue" not in output_layout.columns:
        output_layout["HardcodedValue"] = ""

    return {
        "ProductRoute": product_route,
        "OptionRules": option_rules,
        "OutputLayout": output_layout,
        "_debug_OptionRules_raw_headers": _debug_option_raw_headers,
        "_debug_OptionRules_renamed_headers": _debug_option_renamed_headers,
    }


# ============== Load Order File ==============

def _get_excel_bytes(uploaded_file, password=None):
    raw = uploaded_file.read()
    if not raw:
        raise ValueError("파일 내용이 비어 있습니다.")
    if not HAS_MSOFFCRYPTO:
        return raw
    bio = BytesIO(raw)
    office_file = msoffcrypto.OfficeFile(bio)
    if not office_file.is_encrypted():
        return raw
    if not (password and str(password).strip()):
        raise ValueError("엑셀 파일이 암호화되어 있습니다. 비밀번호를 입력하세요.")
    try:
        office_file.load_key(password=str(password).strip())
        decrypted = BytesIO()
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted.getvalue()
    except Exception as e:
        err_msg = str(e).lower()
        if "invalidkey" in type(e).__name__.lower() or "password" in err_msg or "decrypt" in err_msg:
            raise ValueError("비밀번호가 올바르지 않습니다.") from e
        raise ValueError(f"암호 해제 실패: {e}") from e


def load_excel(uploaded_file, password=None):
    file_name = (uploaded_file.name or "").lower()
    if not file_name.endswith(".xlsx"):
        raise ValueError("엑셀 파일(.xlsx)이 아닙니다.")
    raw_bytes = _get_excel_bytes(uploaded_file, password=password)
    preview = pd.read_excel(BytesIO(raw_bytes), header=None, nrows=20)
    header_idx = find_header_row(preview)
    df = pd.read_excel(BytesIO(raw_bytes), header=header_idx)
    _strip_columns(df)
    ensure_quantity_column(df)
    return df


def read_csv_with_encoding(file):
    encodings = ("utf-8-sig", "utf-8", "cp949", "euc-kr")
    last_err = None
    for enc in encodings:
        try:
            file.seek(0)
            preview = pd.read_csv(file, encoding=enc, header=None, nrows=20)
            file.seek(0)
            header_idx = find_header_row(preview)
            file.seek(0)
            df = pd.read_csv(file, encoding=enc, header=header_idx)
            _strip_columns(df)
            ensure_quantity_column(df)
            return df
        except Exception as e:
            last_err = e
    raise ValueError(f"CSV 인코딩 판별 실패. {last_err}")


# ============== Step B: Routing ==============

def route_vendor(df, product_route):
    product_route = product_route.sort_values("Priority", ascending=True).reset_index(drop=True)
    name_col = "상품명" if "상품명" in df.columns else None
    option_col = "옵션정보" if "옵션정보" in df.columns else None
    if not name_col:
        df = df.copy()
        df["_VendorID"] = "Unclassified"
        return df

    def search_vendor(row):
        name = str(row.get(name_col, "") or "")
        option = str(row.get(option_col, "") or "") if option_col else ""
        search_text = (name + " " + option).strip()
        fallback_vendor = None
        for _, r in product_route.iterrows():
            keywords_raw = str(r.get("Keyword", "") or "").strip()
            if not keywords_raw:
                continue
            keywords = [k.strip() for k in keywords_raw.split(",") if k.strip()]
            for kw in keywords:
                if str(kw).upper() == "DEFAULT":
                    fallback_vendor = str(r.get("TargetVendorID", "") or "").strip()
                    break
                if kw in search_text:
                    return str(r.get("TargetVendorID", "") or "").strip()
        return fallback_vendor if fallback_vendor else "Unclassified"

    df = df.copy()
    df["_VendorID"] = df.apply(search_vendor, axis=1)
    return df


# ============== Step C: Option Rules (Logic Engine) ==============

def _safe_int(x, default=1):
    try:
        v = int(x)
        return v if v > 0 else default
    except (TypeError, ValueError):
        return default


def _apply_remove_text(text, param):
    if not param or not str(param).strip():
        return text
    keywords = [k.strip() for k in str(param).split(",") if k.strip()]
    for kw in keywords:
        text = str(text).replace(kw, "")
    return text


def _apply_remove_regex(text, param):
    if not param or not str(param).strip():
        return text
    try:
        return re.sub(str(param).strip(), "", str(text))
    except re.error:
        return text


def _apply_mask_text(text, param):
    if not param or not str(param).strip():
        return text
    t = str(text)
    keywords = [k.strip() for k in str(param).split(",") if k.strip()]
    for kw in keywords:
        t = t.replace(kw, f"__MASK__{kw}__")
    return t


def _apply_unmask_text(text, param):
    return re.sub(r"__MASK__(.+?)__", r"\1", str(text))


def _apply_convert_weight_range_fix(text, qty, calculated_weight_ref):
    """
    One-Shot Safe: Find weight patterns. If range (e.g. 800g-1kg), take MAX only.
    weight_kg = MaxValue_kg * RowQty. Update calculated_weight_ref; remove weight text from string.
    Do NOT append weight to text; merge_orders will append the total once.
    Caller must check _weight_calculated lock before invoking (prevents double count).
    """
    qty = _safe_int(qty, 1)
    pattern = re.compile(r"(\d+(?:\.\d+)?)\s*(g|kg|G|KG)", re.IGNORECASE)
    matches = list(pattern.finditer(str(text)))
    if not matches:
        return text
    values_kg = []
    for m in matches:
        num = float(m.group(1))
        u = (m.group(2) or "g").lower()
        kg = num / 1000.0 if u == "g" else num
        values_kg.append(kg)
    max_kg = max(values_kg)
    weight_kg = max_kg * qty
    calculated_weight_ref[0] += weight_kg
    cleaned = pattern.sub("", str(text)).strip()
    return cleaned


def _apply_calc_unit(text, param, qty):
    """
    Precision Mode: Find (\\d+)\\s*{Parameter}, replace with NewNum = FoundNum * RowQty, keep unit.
    Example: "10마리" (Qty 3) -> "30마리".
    Numbers in a range (prefix - or ~) or approximation (suffix 내외/내외)) are left unchanged.
    """
    qty = _safe_int(qty, 1)
    unit = str(param).strip() if param else ""
    if not unit:
        return text
    pattern = re.compile(r"([-~]\s*)?(\d+)\s*" + re.escape(unit) + r"(\s*내외\)?)?")
    def repl(m):
        prefix = m.group(1) or ""
        num_str = m.group(2)
        suffix = m.group(3) or ""
        if "-" in prefix or "~" in prefix or "내외" in suffix:
            return m.group(0)
        n = int(num_str) * qty
        return f"{prefix}{n}{unit}{suffix}"
    return pattern.sub(repl, str(text))


def _apply_group_multiply(text, param, qty, search_in=None):
    """
    If Parameter exists in search_in (or text if search_in not provided), append (x{RowQty}) to text only.
    search_in: optional string to check for param (e.g. product+option); when set, append is still applied to text.
    Example: text="1팩", search_in="Squid 1팩", param="Squid" -> "1팩 (x2)".
    """
    qty = _safe_int(qty, 1)
    kw = str(param).strip() if param else ""
    if not kw:
        return text
    check_against = str(search_in).strip() if search_in is not None else str(text)
    if kw not in check_against:
        return text
    return (str(text).strip() + f" (x{qty})").strip()


def _apply_append_suffix(text, param):
    if not param:
        return text
    return (str(text).strip() + " " + str(param).strip()).strip()


def _apply_prepend_text(text, param):
    """Prepend Parameter to text. Always apply (no lock)."""
    if not param or not str(param).strip():
        return text
    return (str(param).strip() + " " + str(text).strip()).strip()


def _apply_append_qty_unit(text, param, qty):
    """
    Direct Append (Squid Logic): Ignore weight/content, append "Qty + Unit".
    Example: Qty=3, Parameter="팩" -> Appends " 3팩". Result: "Squid 3팩".
    """
    qty = _safe_int(qty, 1)
    unit = str(param).strip() if param else "개"
    return (str(text).strip() + f" {qty}{unit}").strip()


def _apply_format_qty_single_stamp(text, param, qty, qty_display_lock_ref):
    """
    Single Stamp: If not qty_display_lock, append format (e.g. x{qty}개), set lock = True.
    """
    if qty_display_lock_ref[0]:
        return text
    qty = _safe_int(qty, 1)
    if not param or not str(param).strip():
        fmt = f" x{qty}개"
    else:
        fmt = str(param).strip().replace("{qty}", str(qty))
        if "{qty}" not in str(param):
            fmt = f" x{qty}개"
    qty_display_lock_ref[0] = True
    return (str(text).strip() + " " + fmt).strip()


def _apply_replace_regex_sub(text, param):
    """
    REPLACE_REGEX_SUB: Split Parameter by '///' -> pattern, replacement.
    Apply re.sub(pattern, replacement, text).
    Example: '^.*Octopus.*$ /// (Steamed)' replaces matching line with '(Steamed)'.
    """
    if not param or not str(param).strip():
        return text
    s = str(param).strip()
    if "///" in s:
        parts = s.split("///", 1)
        pattern, repl = parts[0].strip(), parts[1].strip()
    elif "||" in s:
        pattern, repl = s.split("||", 1)[0].strip(), s.split("||", 1)[1].strip()
    else:
        pattern, repl = s, ""
    try:
        return re.sub(pattern, repl, str(text))
    except re.error:
        return text


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
            text = _apply_calc_unit(text, param, qty)
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
            text = _apply_group_multiply(text, param, qty, search_in=full_context)
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


# ============== Step D: Merge & Sort ==============

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
        if total_weight == int(total_weight):
            return f"{int(total_weight)}kg"
        return f"{total_weight}kg"

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
            row_qty = _safe_int(row.get("수량"), 1) if has_qty_col else 1
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


# ============== Step E: Export ==============

def build_output_dataframe(merged_df, output_layout, vendor_id):
    layout = output_layout[output_layout["VendorID"].astype(str).str.strip() == str(vendor_id).strip()]
    if layout.empty:
        return None
    layout = layout.sort_values("ExcelCol").reset_index(drop=True)
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
            if pd.notna(src) and str(src).strip() and str(src).strip() in merged_df.columns:
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


def process_all_data(df, config):
    """
    Full pipeline: filter -> route -> option rules -> [SORT by OrderNo] -> merge -> sort by date -> export.
    """
    product_route = config["ProductRoute"]
    option_rules = config["OptionRules"]
    
    # 1. 필터링 (Filter)
    df = filter_instruction_rows(df)
    if df.empty:
        return []
    if "구매자명" not in df.columns:
        df = df.copy()
        df["구매자명"] = ""
        
    # 2. 업체 분류 (Routing)
    df = route_vendor(df, product_route)
    
    # 3. 룰 적용 (Option Engine)
    df = run_option_engine(df, option_rules, debug_log=None)
    
    # =======================================================
    # ⭐⭐ [NEW] 병합 전 정렬 (Pre-Merge Sort) ⭐⭐
    # 설명: '상품주문번호'가 있으면 오름차순(1, 2, 3...)으로 정렬합니다.
    # 이렇게 하면 뒤에 merge_orders가 실행될 때 이 순서대로 옵션이 합쳐집니다.
    # =======================================================
    sort_col = "상품주문번호"
    if sort_col in df.columns:
        # 데이터 타입을 문자열로 통일해서 정렬 (숫자와 문자가 섞여도 에러 안 나게)
        df[sort_col] = df[sort_col].astype(str)
        df = df.sort_values(by=sort_col, ascending=True)
    # =======================================================

    # 4. 병합 (Merge) - 이제 정렬된 순서대로 합쳐짐!
    merged = merge_orders(df, option_rules=option_rules)
    
    # 5. 결제일 기준 최종 정렬 (Sort by Payment Date)
    if "결제일" in merged.columns:
        merged = sort_by_payment_date(merged)
        
    # 6. 파일 생성 (Export)
    return export_individual_files(merged, config)


# ============== UI (v14.3: Session State) ==============

def main():
    st.set_page_config(page_title="속초 발주 처리 시스템 v14.4", layout="wide")
    st.title("속초 발주 처리 시스템 v14.4")

    # Session state: persist processed results so download buttons do NOT trigger re-run
    if "processed_results" not in st.session_state:
        st.session_state.processed_results = None
    if "last_file_id" not in st.session_state:
        st.session_state.last_file_id = None

    st.write("설정: config.xlsx (로컬). 주문 파일 업로드 후 Process 버튼을 클릭하세요.")

    if not HAS_MSOFFCRYPTO:
        st.warning("비밀번호 보호 엑셀: `pip install msoffcrypto-tool`")

    config_path = "config.xlsx"
    try:
        path = Path(config_path).resolve()
        if not path.exists():
            st.error(f"설정 파일을 찾을 수 없습니다. 경로: {path!s}")
            st.info("config.xlsx를 앱과 같은 폴더에 두거나 경로를 확인하세요.")
            return
        cache_key = f"{path.stat().st_mtime}_{path.stat().st_size}"
        config = load_config_local(config_path, _password="1111", cache_key=cache_key)
    except FileNotFoundError as e:
        st.error(str(e))
        return
    except ValueError as e:
        st.error(str(e))
        return

    with st.sidebar:
        st.subheader("Config")
        st.caption("Config loaded from config.xlsx")

    uploaded_file = st.file_uploader("주문 파일 (.xlsx 또는 .csv)", type=["xlsx", "csv"], key="uploaded_file")

    # When a new order file is uploaded (different name/size), reset processed_results
    if uploaded_file is not None:
        file_id = (uploaded_file.name, uploaded_file.size)
        if st.session_state.last_file_id != file_id:
            st.session_state.processed_results = None
            st.session_state.last_file_id = file_id

    # No file and no cached results -> ask to upload
    if uploaded_file is None and st.session_state.processed_results is None:
        st.info("주문 파일을 업로드한 뒤 'Process Orders' 버튼을 클릭하세요.")
        return

    # No file but we have results (e.g. after download click rerun) -> show download only
    if uploaded_file is None and st.session_state.processed_results is not None:
        st.success(f"처리 완료. 총 {len(st.session_state.processed_results)}개 업체 파일을 다운로드하세요.")
        for i, pf in enumerate(st.session_state.processed_results):
            st.download_button(
                label=f"📥 Download [{pf['vendor']}] File",
                data=pf["data"].getvalue(),
                file_name=pf["filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{pf['vendor']}_{i}",
            )
        return

    # File uploaded: load and validate
    password = None
    if (uploaded_file.name or "").lower().endswith(".xlsx"):
        password = st.text_input("주문 엑셀 비밀번호 (없으면 비움)", type="password", key="order_pw")
    file_name = (uploaded_file.name or "").lower()
    try:
        if file_name.endswith(".xlsx"):
            df = load_excel(uploaded_file, password=password)
        else:
            uploaded_file.seek(0)
            df = read_csv_with_encoding(uploaded_file)
    except Exception as e:
        st.error(str(e))
        return

    before = len(df)
    df = filter_instruction_rows(df)
    if before > len(df):
        st.info(f"안내 문구 행 {before - len(df)}개 제거.")
    if df.empty:
        st.warning("처리할 데이터가 없습니다.")
        return

    required = ["수량", "상품명", "옵션정보", "배송비 묶음번호"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"필수 컬럼 누락: {', '.join(str(x) for x in missing)}")
        return

    st.subheader("Raw Data Preview (상위 5행)")
    st.dataframe(df.head())

    # Process button: run pipeline and store in session state
    if st.button("Process Orders"):
        with st.spinner("Processing..."):
            try:
                result = process_all_data(df, config)
                st.session_state.processed_results = result
                st.session_state.last_file_id = (uploaded_file.name, uploaded_file.size)
            except Exception as e:
                st.error(f"처리 실패: {e}")
                return
        st.rerun()

    # If we already have results for this session, show download section
    if st.session_state.processed_results is not None:
        st.success(f"처리 완료. 총 {len(st.session_state.processed_results)}개 업체 파일을 다운로드하세요.")
        for i, pf in enumerate(st.session_state.processed_results):
            st.download_button(
                label=f"📥 Download [{pf['vendor']}] File",
                data=pf["data"].getvalue(),
                file_name=pf["filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{pf['vendor']}_{i}",
            )
    else:
        st.caption("위 'Process Orders' 버튼을 클릭하면 처리 후 다운로드 버튼이 표시됩니다.")


if __name__ == "__main__":
    main()
