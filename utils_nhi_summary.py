# utils_nhi_summary.py (V4 - 最終健壯性修正版)
import pandas as pd
import config
from utils_salary_calc import get_salary_report_for_editing

def get_nhi_employer_summary(conn, year: int):
    """
    計算指定年度中，投保單位(公司)應負擔的二代健保補充保費。
    (V4 - 增加對空月份的處理，避免KeyError)

    Args:
        conn: 資料庫連線。
        year (int): 要計算的年份。

    Returns:
        pd.DataFrame: 包含每月計算明細的報表。
    """
    results = []
    
    # 逐月計算
    for month in range(1, 13):
        month_end_str = f"{year}-{month:02d}-{pd.Timestamp(year, month, 1).days_in_month}"

        # 1. 取得當月薪資報表
        report_df, _ = get_salary_report_for_editing(conn, year, month)

        # 2. [核心修正] 檢查 '應發總額' 欄位是否存在
        # 如果報表為空 (該月無薪資紀錄)，則 paid_salary 為 0
        total_paid_salary = report_df['應發總額'].sum() if '應發總額' in report_df.columns else 0

        # 3. 計算當月「健保投保薪資總額 (B)」
        total_insured_salary = 0
        if not report_df.empty:
            health_grades_query = """
                SELECT MIN(salary_max), MAX(salary_max) FROM insurance_grade 
                WHERE type = 'health' AND start_date = (
                    SELECT MAX(start_date) FROM insurance_grade WHERE type = 'health' AND start_date <= ?
                )
            """
            min_max_grades = conn.cursor().execute(health_grades_query, (month_end_str,)).fetchone()
            min_insured_amount, max_insured_amount = (min_max_grades or (0, 0))
            
            # [核心修正] 同樣檢查 '申報薪資' 欄位是否存在
            if '申報薪資' in report_df.columns:
                for _, emp_row in report_df.iterrows():
                    salary_for_insurance = emp_row['申報薪資']
                    insured_amount = 0

                    if salary_for_insurance > 0:
                        find_grade_query = "SELECT salary_max FROM insurance_grade WHERE type = 'health' AND ? BETWEEN salary_min AND salary_max AND start_date <= ? ORDER BY start_date DESC, grade DESC LIMIT 1"
                        grade_tuple = conn.cursor().execute(find_grade_query, (salary_for_insurance, month_end_str)).fetchone()
                        
                        if grade_tuple:
                            insured_amount = grade_tuple[0]
                        elif max_insured_amount and salary_for_insurance > max_insured_amount:
                            insured_amount = max_insured_amount
                        elif min_insured_amount and salary_for_insurance < min_insured_amount:
                            insured_amount = min_insured_amount

                    total_insured_salary += insured_amount

        # 4. 計算差額與應繳保費
        diff = total_paid_salary - total_insured_salary
        premium = round(diff * config.NHI_SUPPLEMENT_RATE) if diff > 0 else 0
        
        results.append({
            '月份': f"{month}月",
            '支付薪資總額 (A)': total_paid_salary,
            '健保投保薪資總額 (B)': total_insured_salary,
            '計費差額 (A - B)': diff,
            '單位應繳補充保費': premium
        })
        
    summary_df = pd.DataFrame(results)
    total_row = summary_df.sum(numeric_only=True).to_frame().T
    total_row['月份'] = '年度總計'
    summary_df = pd.concat([summary_df, total_row], ignore_index=True)

    return summary_df