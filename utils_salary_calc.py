# utils_salary_calc.py
import pandas as pd
from datetime import datetime
import config
import numpy as np

def get_item_types(conn):
    """獲取所有薪資項目的名稱及其類型 (earning/deduction) 的字典"""
    return pd.read_sql("SELECT name, type FROM salary_item", conn).set_index('name')['type'].to_dict()

def check_salary_records_exist(conn, year, month):
    """檢查指定年月的薪資主紀錄是否存在"""
    return conn.cursor().execute("SELECT 1 FROM salary WHERE year = ? AND month = ? LIMIT 1", (year, month)).fetchone() is not None

def get_active_employees_for_month(conn, year, month):
    """取得在指定年月仍在職的員工，並依ID排序"""
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
    """薪資試算引擎: 純計算，不寫入資料庫，返回一個格式化的 DataFrame"""
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
        
        details = {'員工編號': emp_hr_code, '加保單位': emp_company}

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
        
        total_late_early_minutes = details['遲到(分)'] + details['早退(分)']
        deduction_late_early = int(np.round((total_late_early_minutes / 60) * hourly_rate))
        if deduction_late_early > 0:
            details['遲到'] = -deduction_late_early
            details['早退'] = 0

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
        
        all_salary_data.append({'員工姓名': emp_name, **details})

    if not all_salary_data: return pd.DataFrame(), {}
    return pd.DataFrame(all_salary_data).fillna(0), item_types

def save_salary_df(conn, year, month, df: pd.DataFrame):
    cursor = conn.cursor()
    emp_map = pd.read_sql("SELECT id, name_ch FROM employee", conn).set_index('name_ch')['id'].to_dict()
    item_map = pd.read_sql("SELECT id, name FROM salary_item", conn).set_index('name')['id'].to_dict()
    for _, row in df.iterrows():
        emp_id = emp_map.get(row['員工姓名'])
        if not emp_id: continue
        cursor.execute("INSERT OR IGNORE INTO salary (employee_id, year, month) VALUES (?, ?, ?)", (emp_id, year, month))
        salary_id = cursor.execute("SELECT id FROM salary WHERE employee_id = ? AND year = ? AND month = ?", (emp_id, year, month)).fetchone()[0]
        cursor.execute("DELETE FROM salary_detail WHERE salary_id = ?", (salary_id,))
        details_to_insert = [(salary_id, item_map[k], int(v)) for k, v in row.items() if k in item_map and v != 0]
        if details_to_insert:
            cursor.executemany("INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)", details_to_insert)
    conn.commit()

def get_previous_non_insured_names(conn, current_year, current_month):
    prev_date = (datetime(current_year, current_month, 1) - pd.DateOffset(months=1))
    year, month = prev_date.year, prev_date.month
    item_id_ins_tuple = conn.cursor().execute("SELECT id FROM salary_item WHERE name = '勞健保'").fetchone()
    if not item_id_ins_tuple: return []
    item_id_insurance = item_id_ins_tuple[0]
    query = "SELECT e.name_ch FROM employee e WHERE e.id IN (SELECT s.employee_id FROM salary s LEFT JOIN salary_detail sd ON s.id = sd.salary_id AND sd.salary_item_id = ? WHERE s.year = ? AND s.month = ? GROUP BY s.employee_id HAVING IFNULL(SUM(sd.amount), 0) = 0)"
    names = conn.cursor().execute(query, (item_id_insurance, year, month)).fetchall()
    return [name[0] for name in names]

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

def save_data_editor_changes(conn, year, month, edited_df: pd.DataFrame):
    cursor = conn.cursor()
    emp_map = pd.read_sql("SELECT id, name_ch FROM employee", conn).set_index('name_ch')['id'].to_dict()
    item_map = pd.read_sql("SELECT id, name FROM salary_item", conn).set_index('name')['id'].to_dict()
    for _, row in edited_df.iterrows():
        emp_id = emp_map.get(row.get('員工姓名'))
        if not emp_id: continue
        salary_id_tuple = cursor.execute("SELECT id FROM salary WHERE employee_id = ? AND year = ? AND month = ?", (emp_id, year, month)).fetchone()
        if not salary_id_tuple: continue
        salary_id = salary_id_tuple[0]
        for item_name, amount in row.items():
            item_id = item_map.get(item_name)
            if not item_id or pd.isna(amount): continue
            cursor.execute("DELETE FROM salary_detail WHERE salary_id = ? AND salary_item_id = ?", (salary_id, item_id))
            if amount != 0:
                cursor.execute("INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)", (salary_id, item_id, int(amount)))
        if '匯入銀行' in row and pd.notna(row['匯入銀行']):
            override_amount = int(row['匯入銀行'])
            cursor.execute("UPDATE salary SET bank_transfer_override = ? WHERE id = ?", (override_amount, salary_id))
    conn.commit()

def get_salary_report_for_editing(conn, year, month):
    """讀取已儲存的薪資單，並重新計算公司成本與衍生欄位"""
    month_str = f"{year}-{month:02d}"
    month_last_day = f"{year}-{month:02d}-{pd.Timestamp(year, month, 1).days_in_month}"
    
    emp_info_query = """
    SELECT 
        e.id as employee_id, e.name_ch as '員工姓名', e.hr_code as '員工編號',
        s.bank_transfer_override, c.name as '加保單位',
        sbh.base_salary, sbh.dependents
    FROM salary s
    JOIN employee e ON s.employee_id = e.id
    LEFT JOIN (
        SELECT employee_id, base_salary, dependents, ROW_NUMBER() OVER(PARTITION BY employee_id ORDER BY start_date DESC) as rn
        FROM salary_base_history
    ) sbh ON e.id = sbh.employee_id AND sbh.rn = 1
    LEFT JOIN employee_company_history ech ON e.id = ech.employee_id AND ech.start_date <= ? AND (ech.end_date IS NULL OR ech.end_date >= ?)
    LEFT JOIN company c ON ech.company_id = c.id
    WHERE s.year = ? AND s.month = ?
    GROUP BY e.id
    """
    report_df = pd.read_sql_query(emp_info_query, conn, params=(month_last_day, month_last_day, year, month))
    if report_df.empty: return pd.DataFrame(), {}
    
    details_query = """
    SELECT s.employee_id, si.name as item_name, sd.amount 
    FROM salary_detail sd
    JOIN salary_item si ON sd.salary_item_id = si.id
    JOIN salary s ON sd.salary_id = s.id
    WHERE s.year = ? AND s.month = ?
    """
    details_df = pd.read_sql_query(details_query, conn, params=(year, month))
    
    # [核心修正] 重新計算時數/分鐘數
    attendance_query = f"SELECT employee_id, overtime1_minutes, overtime2_minutes, late_minutes, early_leave_minutes FROM attendance WHERE STRFTIME('%Y-%m', date) = '{month_str}'"
    attendance_df = pd.read_sql_query(attendance_query, conn).groupby('employee_id').sum().reset_index()
    attendance_df['延長工時'] = round(attendance_df.get('overtime1_minutes', 0) / 60, 2)
    attendance_df['再延長工時'] = round(attendance_df.get('overtime2_minutes', 0) / 60, 2)
    attendance_df.rename(columns={'late_minutes': '遲到(分)', 'early_leave_minutes': '早退(分)'}, inplace=True)
    report_df = pd.merge(report_df, attendance_df[['employee_id', '延長工時', '再延長工時', '遲到(分)', '早退(分)']], on='employee_id', how='left')
    
    if not details_df.empty:
        pivot_details = details_df.pivot_table(index='employee_id', columns='item_name', values='amount').fillna(0)
        report_df = pd.merge(report_df, pivot_details, on='employee_id', how='left')

    report_df.fillna(0, inplace=True)
    item_types = get_item_types(conn)
    
    non_insured_names = get_previous_non_insured_names(conn, year, month)
    company_costs = []
    for _, row in report_df.iterrows():
        costs = {}
        base_salary, dependents = row.get('base_salary', 0), row.get('dependents', 0)
        if row['員工姓名'] in non_insured_names:
            costs.update({'公司負擔_勞保': 0, '公司負擔_健保': 0, '勞退提撥(公司負擔)': 0})
        else:
            num_insured = 1 + min(dependents, 3)
            sql_labor_com = "SELECT employer_fee FROM insurance_grade WHERE type = 'labor' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            costs['公司負擔_勞保'] = int((conn.cursor().execute(sql_labor_com, (base_salary,)).fetchone() or [0])[0] or 0)
            sql_health_com = "SELECT employer_fee FROM insurance_grade WHERE type = 'health' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            costs['公司負擔_健保'] = int(((conn.cursor().execute(sql_health_com, (base_salary,)).fetchone() or [0])[0] or 0) * num_insured)
            costs['勞退提撥(公司負擔)'] = int(np.round(base_salary * 0.06))
        company_costs.append(costs)
    
    report_df = pd.concat([report_df, pd.DataFrame(company_costs, index=report_df.index)], axis=1)
    
    earning_cols = [c for c, t in item_types.items() if t == 'earning' and c in report_df.columns]
    deduction_cols = [c for c, t in item_types.items() if t == 'deduction' and c in report_df.columns]
    
    report_df['應發總額'] = report_df[earning_cols].sum(axis=1, numeric_only=True)
    report_df['應扣總額'] = report_df[deduction_cols].sum(axis=1, numeric_only=True)
    report_df['實發淨薪'] = report_df['應發總額'] + report_df['應扣總額']
    
    shenbao_cols = ['底薪', '事假', '病假', '遲到', '早退']
    report_df['申報薪資'] = report_df[[c for c in shenbao_cols if c in report_df.columns]].sum(axis=1)
    bank_cols = ['底薪', '加班費', '勞健保', '事假', '病假', '遲到', '早退']
    report_df['匯入銀行'] = report_df[[c for c in bank_cols if c in report_df.columns]].sum(axis=1)
    
    report_df['匯入銀行'] = report_df['bank_transfer_override'].where(pd.notna(report_df['bank_transfer_override']) & (report_df['bank_transfer_override'] != 0), report_df['匯入銀行'])
    report_df['現金'] = report_df['實發淨薪'] - report_df['匯入銀行']
    
    final_cols = ['員工姓名', '員工編號', '加保單位'] + earning_cols + deduction_cols + ['應發總額', '應扣總額', '實發淨薪', '申報薪資', '匯入銀行', '現金', '公司負擔_勞保', '公司負擔_健保', '勞退提撥(公司負擔)', '延長工時', '再延長工時', '遲到(分)', '早退(分)']
    for col in final_cols:
        if col not in report_df.columns: report_df[col] = 0
            
    result_df = report_df.drop(columns=['employee_id', 'bank_transfer_override', 'base_salary', 'dependents'], errors='ignore')
    result_df = result_df[[c for c in final_cols if c in result_df.columns]]

    for col in result_df.columns:
        if col not in ['員工姓名', '員工編號', '加保單位'] and pd.api.types.is_numeric_dtype(result_df[col]):
            if any(k in col for k in ['(分)', '工時']):
                pass # Keep float for hours
            else:
                result_df[col] = result_df[col].astype(int)
            
    return result_df.sort_values(by='員工編號').reset_index(drop=True), item_types