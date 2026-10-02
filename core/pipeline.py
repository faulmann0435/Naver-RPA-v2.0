"""Full pipeline: process_all_data (filter -> route -> rules -> merge -> sort -> export)."""
from core.engine import run_option_engine
from core.exporter import export_individual_files
from core.merger import filter_instruction_rows, merge_orders, sort_by_payment_date
from core.router import route_vendor


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
