"""Order file loading: header detection, quantity column, xlsx/csv readers."""
from io import BytesIO

import pandas as pd

try:
    import msoffcrypto
    HAS_MSOFFCRYPTO = True
except ImportError:
    HAS_MSOFFCRYPTO = False

# --- Column names (order data) ---
HEADER_KEYWORDS = ["상품명", "수취인명", "옵션정보"]
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
