# utils_salary.py (最終完整修復版)
import pandas as pd
import io
from datetime import datetime

# --- 常數定義 ---
SALARY_ITEM_COLUMNS_MAP = {
    'id': 'ID',
    'name': '項目名稱',
    'type': '類型 (earning/deduction)',
    'is_active': '是否啟用'
}
SALARY_BASE_HISTORY_COLUMNS_MAP = {
    'id': '紀錄ID',
    'employee_id': '員工系統ID',
    'name_ch': '員工姓名',
    'base_salary': '底薪',
    'dependents': '眷屬數',
    'start_date': '生效日',
    'end_date': '結束日',
    'note': '備註'
}

# --- 薪資項目 CRUD ---
def get_all_salary_items(conn, active_only=False):
    """取得所有薪資項目"""
    query = "SELECT * FROM salary_item ORDER BY type, id"
    if active_only:
        query = "SELECT * FROM salary_item WHERE is_active = 1 ORDER BY type, id"
    return pd.read_sql_query(query, conn)

def add_salary_item(conn, data):
    """新增薪資項目，如果不存在的話"""
    cursor = conn.cursor()
    sql = "INSERT OR IGNORE INTO salary_item (name, type, is_active) VALUES (?, ?, ?)"
    cursor.execute(sql, (data['name'], data['type'], data['is_active']))
    conn.commit()

def update_salary_item(conn, item_id, data):
    """更新薪資項目"""
    cursor = conn.cursor()
    sql = "UPDATE salary_item SET name = ?, type = ?, is_active = ? WHERE id = ?"
    cursor.execute(sql, (data['name'], data['type'], data['is_active'], item_id))
    conn.commit()

def delete_salary_item(conn, item_id):
    """刪除薪資項目"""
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON;")
    sql = "DELETE FROM salary_item WHERE id = ?"
    cursor.execute(sql, (item_id,))
    conn.commit()
    return cursor.rowcount

# --- 員工底薪/眷屬異動歷史 CRUD ---
def get_salary_base_history(conn):
    """取得所有員工的底薪/眷屬異動歷史"""
    query = """
    SELECT sh.id, sh.employee_id, e.name_ch, sh.base_salary,
           sh.dependents, sh.start_date, sh.end_date, sh.note
    FROM salary_base_history sh JOIN employee e ON sh.employee_id = e.id
    ORDER BY e.id, sh.start_date DESC
    """
    return pd.read_sql_query(query, conn)

def add_salary_base_history(conn, data):
    """新增一筆底薪/眷屬異動歷史"""
    cursor = conn.cursor()
    sql = "INSERT INTO salary_base_history (employee_id, base_salary, dependents, start_date, end_date, note) VALUES (?, ?, ?, ?, ?, ?)"
    cursor.execute(sql, (
        data['employee_id'], data['base_salary'], data['dependents'],
        data['start_date'], data['end_date'], data['note']
    ))
    conn.commit()

def update_salary_base_history(conn, record_id, data):
    """更新指定的底薪/眷屬異動歷史"""
    cursor = conn.cursor()
    sql = "UPDATE salary_base_history SET base_salary = ?, dependents = ?, start_date = ?, end_date = ?, note = ? WHERE id = ?"
    cursor.execute(sql, (
        data['base_salary'], data['dependents'], data['start_date'],
        data['end_date'], data['note'], record_id
    ))
    conn.commit()

def delete_salary_base_history(conn, record_id):
    """刪除指定的底薪/眷屬異動歷史"""
    cursor = conn.cursor()
    sql = "DELETE FROM salary_base_history WHERE id = ?"
    cursor.execute(sql, (record_id,))
    conn.commit()

# --- 一鍵更新基本工資 ---
def get_employees_below_minimum_wage(conn, new_minimum_wage: int):
    """查詢所有在職且目前底薪低於新基本工資的員工"""
    query = """
    WITH latest_salary AS (
        SELECT employee_id, base_salary, dependents,
               ROW_NUMBER() OVER(PARTITION BY employee_id ORDER BY start_date DESC) as rn
        FROM salary_base_history
    )
    SELECT e.id as employee_id, e.name_ch as "員工姓名",
           ls.base_salary as "目前底薪", ls.dependents as "目前眷屬數"
    FROM employee e
    LEFT JOIN latest_salary ls ON e.id = ls.employee_id
    WHERE e.resign_date IS NULL AND ls.rn = 1 AND ls.base_salary < ?
    ORDER BY e.id;
    """
    return pd.read_sql_query(query, conn, params=(new_minimum_wage,))

def batch_update_basic_salary(conn, preview_df: pd.DataFrame, new_wage: int, effective_date):
    """為預覽列表中的所有員工，批次新增一筆調薪紀錄"""
    cursor = conn.cursor()
    try:
        data_to_insert = [
            (row['employee_id'], new_wage, row['目前眷屬數'],
             effective_date.strftime('%Y-%m-%d'), None,
             f"配合 {effective_date.year} 年基本工資調整")
            for _, row in preview_df.iterrows()
        ]
        sql = "INSERT INTO salary_base_history (employee_id, base_salary, dependents, start_date, end_date, note) VALUES (?, ?, ?, ?, ?, ?)"
        cursor.executemany(sql, data_to_insert)
        conn.commit()
        return cursor.rowcount
    except Exception as e:
        conn.rollback()
        raise e

# --- 薪資單產生與計算核心函式 ---
def get_item_types(conn):
    """獲取所有薪資項目的名稱及其類型 (earning/deduction) 的字典"""
    return pd.read_sql("SELECT name, type FROM salary_item", conn).set_index('name')['type'].to_dict()

def check_salary_records_exist(conn, year, month):
    """檢查指定年月的薪資主紀錄是否存在"""
    query = "SELECT 1 FROM salary WHERE year = ? AND month = ? LIMIT 1"
    return conn.cursor().execute(query, (year, month)).fetchone() is not None

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
    
    # 法規參數設定
    WITHHOLDING_TAX_RATE = 0.05
    WITHHOLDING_TAX_THRESHOLD = 88501
    NHI_SUPPLEMENT_RATE = 0.0211
    NHI_SUPPLEMENT_THRESHOLD = 28590

    employees_df = get_active_employees_for_month(conn, year, month)
    if employees_df.empty: return pd.DataFrame(), {}
    
    item_types = get_item_types(conn)
    all_salary_data = []
    hourly_rate_divisor = 240.0

    for _, emp in employees_df.iterrows():
        emp_id, emp_name = emp['id'], emp['name_ch']
        details = {}

        sql_base = "SELECT base_salary, dependents FROM salary_base_history WHERE employee_id = ? ORDER BY start_date DESC LIMIT 1"
        base_info = conn.cursor().execute(sql_base, (emp_id,)).fetchone()
        base_salary = base_info[0] if base_info and base_info[0] is not None else 0
        dependents = base_info[1] if base_info and base_info[1] is not None else 0.0
        hourly_rate = base_salary / hourly_rate_divisor if hourly_rate_divisor > 0 else 0
        details['底薪'] = base_salary

        if emp_name not in non_insured_names:
            sql_labor = "SELECT employee_fee FROM insurance_grade WHERE type = 'labor' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            labor_fee = (conn.cursor().execute(sql_labor, (base_salary,)).fetchone() or [0])[0] or 0
            sql_health = "SELECT employee_fee FROM insurance_grade WHERE type = 'health' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            health_fee_per_person = (conn.cursor().execute(sql_health, (base_salary,)).fetchone() or [0])[0] or 0
            num_insured = 1 + min(dependents, 3)
            health_fee = health_fee_per_person * num_insured
            details['勞健保'] = -(labor_fee + health_fee)
        else:
            details['勞健保'] = 0

        sql_recurring = "SELECT si.name, esi.amount, si.type FROM employee_salary_item esi JOIN salary_item si ON esi.salary_item_id = si.id WHERE esi.employee_id = ?"
        for name, amount, type in conn.cursor().execute(sql_recurring, (emp_id,)).fetchall():
            details[name] = details.get(name, 0) + (-abs(amount) if type == 'deduction' else abs(amount))

        month_str = f"{year}-{month:02d}"
        sql_leave = "SELECT leave_type, SUM(duration) FROM leave_record WHERE employee_id = ? AND strftime('%Y-%m', start_date) = ? AND status = '已通過' GROUP BY leave_type"
        for leave_type, hours in conn.cursor().execute(sql_leave, (emp_id, month_str)).fetchall():
            if hours and hours > 0:
                if leave_type == '事假': details['事假'] = details.get('事假', 0) - (hours * hourly_rate)
                elif leave_type == '病假': details['病假'] = details.get('病假', 0) - (hours * hourly_rate * 0.5)

        total_earnings = sum(v for k, v in details.items() if item_types.get(k) == 'earning')
        supplement_base = total_earnings - details.get('底薪', 0)
        if supplement_base >= NHI_SUPPLEMENT_THRESHOLD:
            details['二代健保補充費'] = - (supplement_base * NHI_SUPPLEMENT_RATE)
        if total_earnings >= WITHHOLDING_TAX_THRESHOLD:
            details['稅款'] = - (total_earnings * WITHHOLDING_TAX_RATE)

        all_salary_data.append({'員工姓名': emp_name, **details})

    if not all_salary_data: return pd.DataFrame(), {}
    
    result_df = pd.DataFrame(all_salary_data).fillna(0)
    return result_df, item_types

def save_salary_df(conn, year, month, df: pd.DataFrame):
    """將試算的 DataFrame 寫入資料庫"""
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
    """從資料庫中，取得指定月份的上一個月中，勞健保為0的員工名單"""
    prev_date = (datetime(current_year, current_month, 1) - pd.DateOffset(months=1))
    year, month = prev_date.year, prev_date.month
    item_id_ins_tuple = conn.cursor().execute("SELECT id FROM salary_item WHERE name = '勞健保'").fetchone()
    if not item_id_ins_tuple: return []
    item_id_insurance = item_id_ins_tuple[0]
    query = """
    SELECT e.name_ch FROM employee e WHERE e.id IN (
        SELECT s.employee_id FROM salary s
        LEFT JOIN salary_detail sd ON s.id = sd.salary_id AND sd.salary_item_id = ?
        WHERE s.year = ? AND s.month = ?
        GROUP BY s.employee_id HAVING IFNULL(SUM(sd.amount), 0) = 0
    )
    """
    names = conn.cursor().execute(query, (item_id_insurance, year, month)).fetchall()
    return [name[0] for name in names]

def batch_update_salary_details_from_excel(conn, year, month, uploaded_file):
    """從上傳的Excel檔案批次更新薪資明細，並返回處理報告"""
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
        if not emp_id:
            report["skipped_emp"].append(emp_name); continue
        salary_id = (cursor.execute("SELECT id FROM salary WHERE employee_id = ? AND year = ? AND month = ?", (emp_id, year, month)).fetchone() or [None])[0]
        if not salary_id: continue
        for item_name, amount in row.items():
            if item_name == '員工姓名' or pd.isna(amount): continue
            item_info = item_map.get(item_name)
            if not item_info:
                report["skipped_item"].append(item_name); continue
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
    """將 st.data_editor 編輯後的 DataFrame 儲存回資料庫"""
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
    """讀取已儲存的薪資單進行編輯，並進行完整的計算與格式化"""
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
