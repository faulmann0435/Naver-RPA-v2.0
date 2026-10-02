"""Config loader: reads and normalizes config.xlsx (optionally password-protected)."""
from io import BytesIO
from pathlib import Path

import pandas as pd

from core.loader import HAS_MSOFFCRYPTO, _strip_columns

if HAS_MSOFFCRYPTO:
    import msoffcrypto


def load_config(config_path: str, password: str | None = None) -> dict:
    """
    Load config.xlsx with password support. Uses msoffcrypto if encrypted.
    Returns: dict with keys ['ProductRoute', 'OptionRules', 'OutputLayout']
    """
    path = Path(config_path).resolve()
    if not path.exists():
        raise FileNotFoundError(
            f"설정 파일을 찾을 수 없습니다. 경로: {path!s}\n"
            "config.xlsx를 앱과 같은 폴더에 두거나 경로를 확인하세요."
        )
    password = password if password and str(password).strip() else None

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
