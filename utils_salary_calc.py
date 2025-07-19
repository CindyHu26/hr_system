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
    query = "SELECT id, name_ch FROM employee WHERE (entry_date IS NOT NULL AND entry_date <= ?) AND (resign_date IS NULL OR resign_date >= ?) ORDER BY id ASC"
    return pd.read_sql_query(query, conn, params=(month_last_day, month_first_day))

def calculate_salary_df(conn, year, month, non_insured_names: list = None):
    """薪資試算引擎: 純計算，不寫入資料庫，返回一個格式化的 DataFrame"""
    if non_insured_names is None: non_insured_names = []
    
    employees_df = get_active_employees_for_month(conn, year, month)
    if employees_df.empty: return pd.DataFrame(), {}
    
    item_types = get_item_types(conn)
    all_salary_data, hourly_rate_divisor = [], config.HOURLY_RATE_DIVISOR

    for _, emp in employees_df.iterrows():
        emp_id, emp_name = emp['id'], emp['name_ch']
        is_non_insured = emp_name in non_insured_names
        details = {}

        sql_base = "SELECT base_salary, dependents FROM salary_base_history WHERE employee_id = ? ORDER BY start_date DESC LIMIT 1"
        base_info = conn.cursor().execute(sql_base, (emp_id,)).fetchone()
        base_salary = base_info[0] if base_info and base_info[0] is not None else 0
        dependents = base_info[1] if base_info and base_info[1] is not None else 0.0
        hourly_rate = base_salary / hourly_rate_divisor if hourly_rate_divisor > 0 else 0
        details['底薪'] = base_salary
        
        sql_recurring = "SELECT si.name, esi.amount, si.type FROM employee_salary_item esi JOIN salary_item si ON esi.salary_item_id = si.id WHERE esi.employee_id = ?"
        for name, amount, type in conn.cursor().execute(sql_recurring, (emp_id,)).fetchall():
            details[name] = details.get(name, 0) + (-abs(amount) if type == 'deduction' else abs(amount))

        month_str = f"{year}-{month:02d}"
        sql_leave = "SELECT leave_type, SUM(duration) FROM leave_record WHERE employee_id = ? AND strftime('%Y-%m', start_date) = ? AND status = '已通過' GROUP BY leave_type"
        for leave_type, hours in conn.cursor().execute(sql_leave, (emp_id, month_str)).fetchall():
            if hours and hours > 0:
                if leave_type == '事假': details['事假'] = details.get('事假', 0) - (hours * hourly_rate)
                elif leave_type == '病假': details['病假'] = details.get('病假', 0) - (hours * hourly_rate * 0.5)

        total_earnings = sum(v for k, v in details.items() if item_types.get(k, '') == 'earning')

        if is_non_insured:
            details['勞健保'] = 0
            if total_earnings >= config.NHI_SUPPLEMENT_THRESHOLD:
                supplement_fee = total_earnings * config.NHI_SUPPLEMENT_RATE
                # [FIX 2/2] Use np.ceil directly
                details['二代健保補充費'] = - (np.ceil(supplement_fee))
        else:
            sql_labor = "SELECT employee_fee FROM insurance_grade WHERE type = 'labor' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            labor_fee = (conn.cursor().execute(sql_labor, (base_salary,)).fetchone() or [0])[0] or 0
            sql_health = "SELECT employee_fee FROM insurance_grade WHERE type = 'health' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            health_fee_per_person = (conn.cursor().execute(sql_health, (base_salary,)).fetchone() or [0])[0] or 0
            num_insured = 1 + min(dependents, 3)
            health_fee = health_fee_per_person * num_insured
            details['勞健保'] = -(labor_fee + health_fee)
            
            supplement_base = total_earnings - details.get('底薪', 0)
            if supplement_base >= config.NHI_SUPPLEMENT_THRESHOLD:
                supplement_fee = supplement_base * config.NHI_SUPPLEMENT_RATE
                details['二代健保補充費'] = - (np.ceil(supplement_fee))

        if total_earnings >= config.WITHHOLDING_TAX_THRESHOLD:
            details['稅款'] = - (total_earnings * config.WITHHOLDING_TAX_RATE)

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
                cursor.execute("UPDATE salary_detail SET amount = ? WHERE id = ?", (final_amount, detail_id))
            else:
                cursor.execute("INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)", (salary_id, item_info['id'], final_amount))
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
        salary_id = (cursor.execute("SELECT id FROM salary WHERE employee_id = ? AND year = ? AND month = ?", (emp_id, year, month)).fetchone() or [None])[0]
        if not salary_id: continue
        for item_name, amount in row.items():
            item_id = item_map.get(item_name)
            if not item_id or item_name == '員工姓名' or pd.isna(amount): continue
            cursor.execute("UPDATE salary_detail SET amount = ? WHERE salary_id = ? AND salary_item_id = ?", (amount, salary_id, item_id))
    conn.commit()

def get_salary_report_for_editing(conn, year, month):
    query = "SELECT e.id as employee_id, e.name_ch as '員工姓名', si.name as item_name, sd.amount FROM salary s JOIN employee e ON s.employee_id = e.id JOIN salary_detail sd ON s.id = sd.salary_id JOIN salary_item si ON sd.salary_item_id = si.id WHERE s.year = ? AND s.month = ?"
    df = pd.read_sql_query(query, conn, params=(year, month))
    item_types = get_item_types(conn)
    if df.empty: return pd.DataFrame(), item_types
    
    pivot_df = df.pivot_table(index=['employee_id', '員工姓名'], columns='item_name', values='amount').reset_index().fillna(0)
    pivot_df.sort_values('employee_id', inplace=True)
    pivot_df.drop(columns=['employee_id'], inplace=True)
    
    earning_cols = [c for c, t in item_types.items() if t == 'earning' and c in pivot_df.columns]
    deduction_cols = [c for c, t in item_types.items() if t == 'deduction' and c in pivot_df.columns]
    
    pivot_df['應發總額'] = pivot_df[earning_cols].sum(axis=1, numeric_only=True)
    pivot_df['應扣總額'] = pivot_df[deduction_cols].sum(axis=1, numeric_only=True)
    pivot_df['實發淨薪'] = pivot_df['應發總額'] + pivot_df['應扣總額']
    
    pivot_df['申報薪資'] = (pivot_df.get('底薪', 0) + pivot_df.get('事假', 0) + pivot_df.get('病假', 0) + pivot_df.get('遲到', 0) + pivot_df.get('早退', 0)).fillna(0)
    pivot_df['匯入銀行'] = (pivot_df.get('底薪', 0) + pivot_df.get('加班費', 0) + pivot_df.get('勞健保', 0) + pivot_df.get('事假', 0) + pivot_df.get('病假', 0) + pivot_df.get('遲到', 0) + pivot_df.get('早退', 0)).fillna(0)
    pivot_df['現金'] = pivot_df['實發淨薪'] - pivot_df['匯入銀行']

    total_cols = ['實發淨薪', '應發總額', '應扣總額', '申報薪資', '匯入銀行', '現金']
    final_cols = ['員工姓名'] + [c for c in total_cols if c in pivot_df.columns] + [c for c in earning_cols if c in pivot_df.columns] + [c for c in deduction_cols if c in pivot_df.columns]
    result_df = pivot_df[[c for c in final_cols if c in pivot_df.columns]]

    for col in result_df.columns:
        if col != '員工姓名' and pd.api.types.is_numeric_dtype(result_df[col]):
            result_df[col] = result_df[col].astype(int)
            
    return result_df, item_types