# utils_salary_engine.py
import pandas as pd
from datetime import datetime
import config
import numpy as np
from utils_special_overtime import calculate_special_overtime_pay

def get_active_employees_for_month(conn, year, month):
    month_first_day = f"{year}-{month:02d}-01"
    last_day = pd.Timestamp(year, month, 1).days_in_month
    month_last_day = f"{year}-{month:02d}-{last_day}"
    query = """
    SELECT 
        e.id, e.name_ch, e.hr_code
    FROM employee e
    WHERE 
        (e.entry_date IS NOT NULL AND e.entry_date <= ?) 
        AND (e.resign_date IS NULL OR e.resign_date >= ?)
    ORDER BY e.id ASC
    """
    return pd.read_sql_query(query, conn, params=(month_last_day, month_first_day))

def calculate_salary_df(conn, year, month):
    """
    薪資試算引擎: 從零開始，根據出勤、假單、底薪等計算全新的薪資草稿。
    """
    employees_df = get_active_employees_for_month(conn, year, month)
    if employees_df.empty: return pd.DataFrame(), {}
    
    month_str = f"{year}-{month:02d}"
    attendance_query = f"SELECT employee_id, overtime1_minutes, overtime2_minutes, late_minutes, early_leave_minutes FROM attendance WHERE STRFTIME('%Y-%m', date) = '{month_str}'"
    monthly_attendance_summary = pd.read_sql_query(attendance_query, conn).groupby('employee_id').sum()

    item_types = pd.read_sql("SELECT name, type FROM salary_item", conn).set_index('name')['type'].to_dict()
    all_salary_data = []

    for _, emp in employees_df.iterrows():
        emp_id, emp_name = emp['id'], emp['name_ch']
        details = {'員工姓名': emp_name, '員工編號': emp['hr_code']}
        
        sql_base = "SELECT base_salary, dependents FROM salary_base_history WHERE employee_id = ? AND start_date <= ? ORDER BY start_date DESC LIMIT 1"
        month_end_date = f"{year}-{month:02d}-{pd.Timestamp(year, month, 1).days_in_month}"
        base_info = conn.cursor().execute(sql_base, (emp_id, month_end_date)).fetchone()
        base_salary = base_info[0] if base_info and base_info[0] is not None else 0
        dependents = base_info[1] if base_info and base_info[1] is not None else 0.0
        hourly_rate = base_salary / config.HOURLY_RATE_DIVISOR if config.HOURLY_RATE_DIVISOR > 0 else 0
        details['底薪'] = base_salary
        
        if emp_id in monthly_attendance_summary.index:
            emp_attendance = monthly_attendance_summary.loc[emp_id]
            late_minutes = emp_attendance.get('late_minutes', 0)
            if late_minutes > 0: details['遲到'] = -int(round((late_minutes / 60) * hourly_rate))
            early_leave_minutes = emp_attendance.get('early_leave_minutes', 0)
            if early_leave_minutes > 0: details['早退'] = -int(round((early_leave_minutes / 60) * hourly_rate))
            overtime1_hours = emp_attendance.get('overtime1_minutes', 0) / 60
            if overtime1_hours > 0: details['加班費(平日)'] = int(round(overtime1_hours * hourly_rate * 1.34))
            overtime2_hours = emp_attendance.get('overtime2_minutes', 0) / 60
            if overtime2_hours > 0: details['加班費(假日)'] = int(round(overtime2_hours * hourly_rate * 1.67))

        sql_leave = "SELECT leave_type, SUM(duration) FROM leave_record WHERE employee_id = ? AND strftime('%Y-%m', start_date) = ? AND status = '已通過' GROUP BY leave_type"
        for leave_type, hours in conn.cursor().execute(sql_leave, (emp_id, month_str)).fetchall():
            if hours and hours > 0:
                if leave_type == '事假': details['事假'] = -int(round(hours * hourly_rate))
                elif leave_type == '病假': details['病假'] = -int(round(hours * hourly_rate * 0.5))

        sql_recurring = "SELECT si.name, esi.amount, si.type FROM employee_salary_item esi JOIN salary_item si ON esi.salary_item_id = si.id WHERE esi.employee_id = ?"
        for name, amount, type in conn.cursor().execute(sql_recurring, (emp_id,)).fetchall():
            details[name] = details.get(name, 0) + (-abs(amount) if type == 'deduction' else abs(amount))

        if base_salary > 0:
            sql_labor = "SELECT employee_fee FROM insurance_grade WHERE type = 'labor' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            labor_fee = (conn.cursor().execute(sql_labor, (base_salary,)).fetchone() or [0])[0]
            sql_health = "SELECT employee_fee FROM insurance_grade WHERE type = 'health' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            health_fee_per_person = (conn.cursor().execute(sql_health, (base_salary,)).fetchone() or [0])[0]
            num_insured = 1 + min(dependents, 3)
            total_health_fee = health_fee_per_person * num_insured
            details['勞健保'] = -int(labor_fee + total_health_fee)

        special_overtime_pay = calculate_special_overtime_pay(conn, emp_id, year, month, hourly_rate)
        if special_overtime_pay > 0:
            details['津貼加班'] = special_overtime_pay

        all_salary_data.append(details)

    return pd.DataFrame(all_salary_data).fillna(0), item_types