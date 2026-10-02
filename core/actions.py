"""Option-rule action helpers (_safe_int and every _apply_* function)."""
import re


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
    One-Shot Safe: Find weight patterns outside parentheses only. If range (e.g. 800g-1kg), take MAX only.
    weight_kg = MaxValue_kg * RowQty. Update calculated_weight_ref; remove weight text from string.
    Do NOT append weight to text; merge_orders will append the total once.
    Caller must check _weight_calculated lock before invoking (prevents double count).
    """
    qty = _safe_int(qty, 1)
    weight_pattern = re.compile(r"(\d+(?:\.\d+)?)\s*(g|kg|G|KG)(?!\s*내외)", re.IGNORECASE)
    paren_pattern = re.compile(r"\([^)]*\)")

    # Mask parenthetical content so weights inside parentheses are not converted
    paren_store = {}
    counter = [0]
    def _mask(m):
        key = f"\x00P{counter[0]}\x00"
        paren_store[key] = m.group(0)
        counter[0] += 1
        return key
    masked = paren_pattern.sub(_mask, str(text))

    matches = list(weight_pattern.finditer(masked))
    if not matches:
        return text  # no weights outside parentheses — return original unchanged

    values_kg = []
    for m in matches:
        num = float(m.group(1))
        u = (m.group(2) or "g").lower()
        kg = num / 1000.0 if u == "g" else num
        values_kg.append(kg)
    max_kg = max(values_kg)
    weight_kg = max_kg * qty
    calculated_weight_ref[0] += weight_kg

    cleaned = weight_pattern.sub("", masked).strip()
    for key, val in paren_store.items():
        cleaned = cleaned.replace(key, val)
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
