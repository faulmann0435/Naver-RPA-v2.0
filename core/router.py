"""Vendor routing: route_vendor."""


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
