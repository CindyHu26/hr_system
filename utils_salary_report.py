# utils_salary_report.py
import pandas as pd
from datetime import datetime

def get_salary_report_for_editing(conn, year, month):
    """
    薪資報表產生器: 從資料庫讀取資料，並處理草稿/定版邏輯後呈現。
    """
    month_first_day = f"{year}-{month:02d}-01"
    month_last_day = f"{year}-{month:02d}-{pd.Timestamp(year, month, 1).days_in_month}"
    
    active_emp_query = """
    SELECT id as employee_id, name_ch as '員工姓名', hr_code as '員工編號'
    FROM employee WHERE (entry_date <= ?) AND (resign_date IS NULL OR resign_date >= ?)
    """
    report_df = pd.read_sql_query(active_emp_query, conn, params=(month_last_day, month_first_day))
    if report_df.empty: return pd.DataFrame(), {}

    salary_main_query = "SELECT * FROM salary WHERE year = ? AND month = ?"
    salary_main_df = pd.read_sql_query(salary_main_query, conn, params=(year, month))
    
    details_query = "SELECT s.employee_id, si.name as item_name, sd.amount FROM salary_detail sd JOIN salary_item si ON sd.salary_item_id = si.id JOIN salary s ON sd.salary_id = s.id WHERE s.year = ? AND s.month = ?"
    details_df = pd.read_sql_query(details_query, conn, params=(year, month))
    
    pivot_details = pd.DataFrame()
    if not details_df.empty:
        pivot_details = details_df.pivot_table(index='employee_id', columns='item_name', values='amount')
    
    report_df = pd.merge(report_df, salary_main_df, on='employee_id', how='left')
    if not pivot_details.empty:
        report_df = pd.merge(report_df, pivot_details, on='employee_id', how='left')

    report_df['status'] = report_df['status'].fillna('draft')

    item_types = pd.read_sql("SELECT name, type FROM salary_item", conn).set_index('name')['type'].to_dict()
    numeric_cols = [col for col in item_types.keys() if col in report_df.columns]
    report_df[numeric_cols] = report_df[numeric_cols].fillna(0)
    
    other_numeric_cols = ['total_payable', 'total_deduction', 'net_salary', 'bank_transfer_amount', 'cash_amount']
    for col in other_numeric_cols:
        if col in report_df.columns:
            report_df[col] = report_df[col].fillna(0)

    earning_cols = [c for c, t in item_types.items() if t == 'earning' and c in report_df.columns]
    deduction_cols = [c for c, t in item_types.items() if t == 'deduction' and c in report_df.columns]
    
    draft_mask = report_df['status'] == 'draft'
    if draft_mask.any():
        report_df.loc[draft_mask, '應付總額'] = report_df.loc[draft_mask, earning_cols].sum(axis=1, numeric_only=True)
        report_df.loc[draft_mask, '應扣總額'] = report_df.loc[draft_mask, deduction_cols].sum(axis=1, numeric_only=True)
        report_df.loc[draft_mask, '實發薪資'] = report_df.loc[draft_mask, '應付總額'] + report_df.loc[draft_mask, '應扣總額']
        report_df.loc[draft_mask, '匯入銀行'] = report_df.loc[draft_mask, '實發薪資'] # 草稿預設
        report_df.loc[draft_mask, '現金'] = 0 # 草稿預設

    final_mask = report_df['status'] == 'final'
    if final_mask.any():
        report_df.loc[final_mask, '應付總額'] = report_df.loc[final_mask, 'total_payable']
        report_df.loc[final_mask, '應扣總額'] = report_df.loc[final_mask, 'total_deduction']
        report_df.loc[final_mask, '實發薪資'] = report_df.loc[final_mask, 'net_salary']
        report_df.loc[final_mask, '匯入銀行'] = report_df.loc[final_mask, 'bank_transfer_amount']
        report_df.loc[final_mask, '現金'] = report_df.loc[final_mask, 'cash_amount']

    final_cols = ['員工姓名', '員工編號', 'status'] + list(item_types.keys()) + ['應付總額', '應扣總額', '實發薪資', '匯入銀行', '現金']
    for col in final_cols:
        if col not in report_df.columns:
            report_df[col] = 0

    return report_df[[c for c in final_cols if c in report_df.columns]].sort_values(by='員工編號').reset_index(drop=True), item_types