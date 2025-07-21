# utils_salary_calc.py (基於您提供的版本進行修正)
import pandas as pd
from datetime import datetime
import config
import numpy as np

def get_item_types(conn):
    return pd.read_sql("SELECT name, type FROM salary_item", conn).set_index('name')['type'].to_dict()

def check_salary_records_exist(conn, year, month):
    return conn.cursor().execute("SELECT 1 FROM salary WHERE year = ? AND month = ? LIMIT 1", (year, month)).fetchone() is not None

def get_active_employees_for_month(conn, year, month):
    month_first_day = f"{year}-{month:02d}-01"
    last_day = pd.Timestamp(year, month, 1).days_in_month
    month_last_day = f"{year}-{month:02d}-{last_day}"
    query = """
    SELECT 
        e.id, e.name_ch, e.hr_code, e.nationality, e.arrival_date,
        c.name as company_name
    FROM employee e
    LEFT JOIN employee_company_history ech ON e.id = ech.employee_id AND ech.start_date <= ? AND (ech.end_date IS NULL OR ech.end_date >= ?)
    LEFT JOIN company c ON ech.company_id = c.id
    WHERE 
        (e.entry_date IS NOT NULL AND e.entry_date <= ?) 
        AND (e.resign_date IS NULL OR e.resign_date >= ?)
    GROUP BY e.id
    ORDER BY e.id ASC
    """
    return pd.read_sql_query(query, conn, params=(month_last_day, month_first_day, month_last_day, month_first_day))

def calculate_salary_df(conn, year, month, non_insured_names: list = None):
    if non_insured_names is None: non_insured_names = []
    
    employees_df = get_active_employees_for_month(conn, year, month)
    if employees_df.empty: return pd.DataFrame(), {}
    
    month_str = f"{year}-{month:02d}"
    attendance_query = f"SELECT employee_id, overtime1_minutes, overtime2_minutes, late_minutes, early_leave_minutes FROM attendance WHERE STRFTIME('%Y-%m', date) = '{month_str}'"
    monthly_attendance_summary = pd.read_sql_query(attendance_query, conn).groupby('employee_id').sum()

    item_types = get_item_types(conn)
    all_salary_data = []

    for _, emp in employees_df.iterrows():
        emp_id, emp_name, emp_hr_code = emp['id'], emp['name_ch'], emp.get('hr_code')
        emp_company = emp.get('company_name')
        emp_nationality = emp.get('nationality', 'TW')
        emp_arrival_date = pd.to_datetime(emp.get('arrival_date')) if pd.notna(emp.get('arrival_date')) else None
        is_non_insured = emp_name in non_insured_names
        
        details = {'員工姓名': emp_name, '員工編號': emp_hr_code, '加保單位': emp_company}

        sql_base = "SELECT base_salary, dependents FROM salary_base_history WHERE employee_id = ? ORDER BY start_date DESC LIMIT 1"
        base_info = conn.cursor().execute(sql_base, (emp_id,)).fetchone()
        base_salary = base_info[0] if base_info and base_info[0] is not None else 0
        dependents = base_info[1] if base_info and base_info[1] is not None else 0.0
        hourly_rate = base_salary / config.HOURLY_RATE_DIVISOR if config.HOURLY_RATE_DIVISOR > 0 else 0
        details['底薪'] = base_salary
        
        if emp_id in monthly_attendance_summary.index:
            emp_attendance = monthly_attendance_summary.loc[emp_id]
            details['延長工時'] = round(emp_attendance.get('overtime1_minutes', 0) / 60, 2)
            details['再延長工時'] = round(emp_attendance.get('overtime2_minutes', 0) / 60, 2)
            details['遲到(分)'] = emp_attendance.get('late_minutes', 0)
            details['早退(分)'] = emp_attendance.get('early_leave_minutes', 0)
        else:
            details.update({'延長工時': 0.0, '再延長工時': 0.0, '遲到(分)': 0, '早退(分)': 0})

        details['加班費'] = int(np.round(details['延長工時'] * hourly_rate * 1.34))
        details['加班費2'] = int(np.round(details['再延長工時'] * hourly_rate * 1.67))
        
        # [核心修正] 將遲到與早退的扣款分開計算
        late_minutes = details['遲到(分)']
        if late_minutes > 0:
            details['遲到'] = -int(np.round((late_minutes / 60) * hourly_rate))

        early_leave_minutes = details['早退(分)']
        if early_leave_minutes > 0:
            details['早退'] = -int(np.round((early_leave_minutes / 60) * hourly_rate))

        sql_leave = "SELECT leave_type, SUM(duration) FROM leave_record WHERE employee_id = ? AND strftime('%Y-%m', start_date) = ? AND status = '已通過' GROUP BY leave_type"
        for leave_type, hours in conn.cursor().execute(sql_leave, (emp_id, month_str)).fetchall():
            if hours and hours > 0:
                if leave_type == '事假': details['事假'] = -int(np.round(hours * hourly_rate))
                elif leave_type == '病假': details['病假'] = -int(np.round(hours * hourly_rate * 0.5))

        sql_recurring = "SELECT si.name, esi.amount, si.type FROM employee_salary_item esi JOIN salary_item si ON esi.salary_item_id = si.id WHERE esi.employee_id = ?"
        for name, amount, type in conn.cursor().execute(sql_recurring, (emp_id,)).fetchall():
            details[name] = details.get(name, 0) + (-abs(amount) if type == 'deduction' else abs(amount))

        total_earnings = sum(v for k, v in details.items() if item_types.get(k, '') == 'earning')
        
        if is_non_insured:
            details.update({'勞健保': 0, '公司負擔_勞保': 0, '公司負擔_健保': 0, '勞退提撥(公司負擔)': 0})
            if total_earnings >= config.NHI_SUPPLEMENT_THRESHOLD:
                details['二代健保補充費'] = -int(np.ceil(total_earnings * config.NHI_SUPPLEMENT_RATE))
        else:
            sql_labor_emp = "SELECT employee_fee, employer_fee FROM insurance_grade WHERE type = 'labor' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            labor_fees = conn.cursor().execute(sql_labor_emp, (base_salary,)).fetchone() or (0, 0)
            sql_health_emp = "SELECT employee_fee, employer_fee FROM insurance_grade WHERE type = 'health' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            health_fees_person = conn.cursor().execute(sql_health_emp, (base_salary,)).fetchone() or (0, 0)
            
            num_insured = 1 + min(dependents, 3)
            details['勞健保'] = -int(labor_fees[0] + (health_fees_person[0] * num_insured))
            details['公司負擔_勞保'] = int(labor_fees[1])
            details['公司負擔_健保'] = int(health_fees_person[1] * num_insured)
            details['勞退提撥(公司負擔)'] = int(np.round(base_salary * 0.06))

        tax_amount = 0
        is_foreigner = emp_nationality and emp_nationality.upper() != 'TW'
        if is_foreigner:
            is_resident = (datetime(year, month, 1) - emp_arrival_date).days >= 183 if emp_arrival_date else False
            if is_resident:
                if total_earnings >= config.WITHHOLDING_TAX_THRESHOLD: tax_amount = total_earnings * config.WITHHOLDING_TAX_RATE
            else:
                threshold = config.NHI_SUPPLEMENT_THRESHOLD * config.FOREIGNER_TAX_RATE_THRESHOLD_MULTIPLIER
                tax_rate = config.FOREIGNER_HIGH_INCOME_TAX_RATE if total_earnings > threshold else config.FOREIGNER_LOW_INCOME_TAX_RATE
                tax_amount = total_earnings * tax_rate
        else:
            if total_earnings >= config.WITHHOLDING_TAX_THRESHOLD: tax_amount = total_earnings * config.WITHHOLDING_TAX_RATE
        
        if tax_amount > 0: details['稅款'] = -int(np.round(tax_amount))
        
        all_salary_data.append(details)

    return pd.DataFrame(all_salary_data).fillna(0), item_types

# --- [功能恢復] 以下是之前版本中遺失的所有函式 ---

def save_salary_draft(conn, year, month, df: pd.DataFrame):
    cursor = conn.cursor()
    emp_map = pd.read_sql("SELECT id, name_ch FROM employee", conn).set_index('name_ch')['id'].to_dict()
    item_map = pd.read_sql("SELECT id, name FROM salary_item", conn).set_index('name')['id'].to_dict()
    for _, row in df.iterrows():
        emp_id = emp_map.get(row['員工姓名'])
        if not emp_id: continue
        cursor.execute("INSERT INTO salary (employee_id, year, month, status) VALUES (?, ?, ?, 'draft') ON CONFLICT(employee_id, year, month) DO UPDATE SET status = 'draft' WHERE status != 'final'", (emp_id, year, month))
        salary_id = cursor.execute("SELECT id FROM salary WHERE employee_id = ? AND year = ? AND month = ?", (emp_id, year, month)).fetchone()[0]
        cursor.execute("DELETE FROM salary_detail WHERE salary_id = ?", (salary_id,))
        details_to_insert = [(salary_id, item_map.get(k), int(v)) for k, v in row.items() if item_map.get(k) and v != 0]
        if details_to_insert:
            cursor.executemany("INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)", details_to_insert)
    conn.commit()

def finalize_salary_records(conn, year, month, df: pd.DataFrame):
    cursor = conn.cursor()
    emp_map = pd.read_sql("SELECT id, name_ch FROM employee", conn).set_index('name_ch')['id'].to_dict()
    for _, row in df.iterrows():
        emp_id = emp_map.get(row['員工姓名'])
        if not emp_id: continue
        params = {
            'total_payable': row.get('應付總額', 0), 'total_deduction': row.get('應扣總額', 0),
            'net_salary': row.get('實發薪資', 0), 'bank_transfer_amount': row.get('匯入銀行', 0),
            'cash_amount': row.get('現金', 0), 'status': 'final',
            'employee_id': emp_id, 'year': year, 'month': month
        }
        cursor.execute("""
            UPDATE salary SET
            total_payable = :total_payable, total_deduction = :total_deduction,
            net_salary = :net_salary, bank_transfer_amount = :bank_transfer_amount,
            cash_amount = :cash_amount, status = :status
            WHERE employee_id = :employee_id AND year = :year AND month = :month
        """, params)
    conn.commit()

def revert_salary_to_draft(conn, year, month, employee_ids: list):
    if not employee_ids: return 0
    cursor = conn.cursor()
    placeholders = ','.join('?' for _ in employee_ids)
    sql = f"UPDATE salary SET status = 'draft' WHERE year = ? AND month = ? AND employee_id IN ({placeholders}) AND status = 'final'"
    params = [year, month] + employee_ids
    cursor.execute(sql, params)
    conn.commit()
    return cursor.rowcount

def batch_update_salary_details_from_excel(conn, year, month, uploaded_file):
    report = {"success": [], "skipped_emp": [], "skipped_item": []}
    df = pd.read_excel(uploaded_file)
    cursor = conn.cursor()
    emp_map = pd.read_sql("SELECT id, name_ch FROM employee", conn).set_index('name_ch')['id'].to_dict()
    item_map_df = pd.read_sql("SELECT id, name, type FROM salary_item", conn)
    item_map = {row['name']: {'id': row['id'], 'type': row['type']} for _, row in item_map_df.iterrows()}
    for _, row in df.iterrows():
        emp_name = row.get('員工姓名')
        if not emp_name: continue
        emp_id = emp_map.get(emp_name)
        if not emp_id: report["skipped_emp"].append(emp_name); continue
        salary_id = (cursor.execute("SELECT id FROM salary WHERE employee_id = ? AND year = ? AND month = ?", (emp_id, year, month)).fetchone() or [None])[0]
        if not salary_id: continue
        for item_name, amount in row.items():
            if item_name == '員工姓名' or pd.isna(amount): continue
            item_info = item_map.get(item_name)
            if not item_info: report["skipped_item"].append(item_name); continue
            final_amount = -abs(float(amount)) if item_info['type'] == 'deduction' else abs(float(amount))
            detail_id = (cursor.execute("SELECT id FROM salary_detail WHERE salary_id = ? AND salary_item_id = ?", (salary_id, item_info['id'])).fetchone() or [None])[0]
            if detail_id:
                cursor.execute("UPDATE salary_detail SET amount = ? WHERE id = ?", (int(final_amount), detail_id))
            else:
                cursor.execute("INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)", (salary_id, item_info['id'], int(final_amount)))
            report["success"].append(f"{emp_name} - {item_name}")
    conn.commit()
    report["skipped_emp"] = list(set(report["skipped_emp"]))
    report["skipped_item"] = list(set(report["skipped_item"]))
    return report

def get_salary_report_for_editing(conn, year, month):
    # This function now correctly combines data for display
    active_emp_df = get_active_employees_for_month(conn, year, month)
    if active_emp_df.empty: return pd.DataFrame(), {}
    
    report_df = active_emp_df.rename(columns={'id': 'employee_id'})
    
    salary_main_query = "SELECT * FROM salary WHERE year = ? AND month = ?"
    salary_main_df = pd.read_sql_query(salary_main_query, conn, params=(year, month))
    details_query = "SELECT s.employee_id, si.name as item_name, sd.amount FROM salary_detail sd JOIN salary_item si ON sd.salary_item_id = si.id JOIN salary s ON sd.salary_id = s.id WHERE s.year = ? AND s.month = ?"
    details_df = pd.read_sql_query(details_query, conn, params=(year, month))
    pivot_details = details_df.pivot_table(index='employee_id', columns='item_name', values='amount')
    
    report_df = pd.merge(report_df, salary_main_df, on='employee_id', how='left')
    if not pivot_details.empty:
        report_df = pd.merge(report_df, pivot_details, on='employee_id', how='left')

    report_df['status'] = report_df['status'].fillna('draft')
    report_df.fillna(0, inplace=True)
    
    item_types = get_item_types(conn)
    earning_cols = [c for c, t in item_types.items() if t == 'earning' and c in report_df.columns]
    deduction_cols = [c for c, t in item_types.items() if t == 'deduction' and c in report_df.columns]
    
    draft_mask = report_df['status'] == 'draft'
    if draft_mask.any():
        report_df.loc[draft_mask, '應付總額'] = report_df.loc[draft_mask, earning_cols].sum(axis=1, numeric_only=True)
        report_df.loc[draft_mask, '應扣總額'] = report_df.loc[draft_mask, deduction_cols].sum(axis=1, numeric_only=True)
        report_df.loc[draft_mask, '實發薪資'] = report_df.loc[draft_mask, '應付總額'] + report_df.loc[draft_mask, '應扣總額']
    
    final_mask = report_df['status'] == 'final'
    if final_mask.any():
        report_df.loc[final_mask, '應付總額'] = report_df.loc[final_mask, 'total_payable']
        report_df.loc[final_mask, '應扣總額'] = report_df.loc[final_mask, 'total_deduction']
        report_df.loc[final_mask, '實發薪資'] = report_df.loc[final_mask, 'net_salary']
        report_df.loc[final_mask, '匯入銀行'] = report_df.loc[final_mask, 'bank_transfer_amount']
        report_df.loc[final_mask, '現金'] = report_df.loc[final_mask, 'cash_amount']

    report_df.loc[draft_mask, '匯入銀行'] = report_df.loc[draft_mask, '實發薪資']
    report_df['現金'] = report_df['實發薪資'] - report_df['匯入銀行']
            
    final_cols = ['員工姓名', '員工編號', 'status'] + earning_cols + deduction_cols + ['應付總額', '應扣總額', '實發薪資', '匯入銀行', '現金']
    for col in final_cols:
        if col not in report_df.columns:
            report_df[col] = 0

    return report_df[final_cols].sort_values(by='員工編號').reset_index(drop=True), item_types