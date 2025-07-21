# utils_salary_crud.py
import pandas as pd

# --- 常數定義 ---
SALARY_ITEM_COLUMNS_MAP = {
    'id': 'ID',
    'name': '項目名稱',
    'type': '類型 (給付earning/扣除deduction)',
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
    query = "SELECT * FROM salary_item ORDER BY type, id"
    if active_only:
        query = "SELECT * FROM salary_item WHERE is_active = 1 ORDER BY type, id"
    return pd.read_sql_query(query, conn)

def add_salary_item(conn, data):
    cursor = conn.cursor()
    sql = "INSERT OR IGNORE INTO salary_item (name, type, is_active) VALUES (?, ?, ?)"
    cursor.execute(sql, (data['name'], data['type'], data['is_active']))
    conn.commit()

def update_salary_item(conn, item_id, data):
    cursor = conn.cursor()
    sql = "UPDATE salary_item SET name = ?, type = ?, is_active = ? WHERE id = ?"
    cursor.execute(sql, (data['name'], data['type'], data['is_active'], item_id))
    conn.commit()

def delete_salary_item(conn, item_id):
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON;")
    sql = "DELETE FROM salary_item WHERE id = ?"
    cursor.execute(sql, (item_id,))
    conn.commit()
    return cursor.rowcount

# --- 員工底薪/眷屬異動歷史 CRUD ---
def get_salary_base_history(conn):
    query = """
    SELECT sh.id, sh.employee_id, e.name_ch, sh.base_salary,
           sh.dependents, sh.start_date, sh.end_date, sh.note
    FROM salary_base_history sh JOIN employee e ON sh.employee_id = e.id
    ORDER BY e.id, sh.start_date DESC
    """
    return pd.read_sql_query(query, conn)

def add_salary_base_history(conn, data):
    cursor = conn.cursor()
    sql = "INSERT INTO salary_base_history (employee_id, base_salary, dependents, start_date, end_date, note) VALUES (?, ?, ?, ?, ?, ?)"
    cursor.execute(sql, (data['employee_id'], data['base_salary'], data['dependents'], data['start_date'], data['end_date'], data['note']))
    conn.commit()

def update_salary_base_history(conn, record_id, data):
    cursor = conn.cursor()
    sql = "UPDATE salary_base_history SET base_salary = ?, dependents = ?, start_date = ?, end_date = ?, note = ? WHERE id = ?"
    cursor.execute(sql, (data['base_salary'], data['dependents'], data['start_date'], data['end_date'], data['note'], record_id))
    conn.commit()

def delete_salary_base_history(conn, record_id):
    cursor = conn.cursor()
    sql = "DELETE FROM salary_base_history WHERE id = ?"
    cursor.execute(sql, (record_id,))
    conn.commit()

# --- 一鍵更新基本工資 ---
def get_employees_below_minimum_wage(conn, new_minimum_wage: int):
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

# --- 員工常態薪資項設定 ---
def get_employee_salary_items(conn):
    query = "SELECT esi.id, e.id as employee_id, e.name_ch as '員工姓名', si.id as salary_item_id, si.name as '項目名稱', si.type as '類型', esi.amount as '金額', esi.start_date as '生效日', esi.end_date as '結束日', esi.note as '備註' FROM employee_salary_item esi JOIN employee e ON esi.employee_id = e.id JOIN salary_item si ON esi.salary_item_id = si.id ORDER BY e.name_ch, si.name"
    return pd.read_sql_query(query, conn)

def get_settings_grouped_by_amount(conn, salary_item_id):
    if not salary_item_id: return {}
    query = "SELECT esi.amount, e.id as employee_id, e.name_ch FROM employee_salary_item esi JOIN employee e ON esi.employee_id = e.id WHERE esi.salary_item_id = ? ORDER BY esi.amount, e.name_ch"
    df = pd.read_sql_query(query, conn, params=(int(salary_item_id),))
    return {amount: group[['employee_id', 'name_ch']].to_dict('records') for amount, group in df.groupby('amount')} if not df.empty else {}

def batch_add_employee_salary_items(conn, employee_ids, salary_item_id, amount, start_date, end_date, note):
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN TRANSACTION")
        placeholders = ','.join('?' for _ in employee_ids)
        cursor.execute(f"DELETE FROM employee_salary_item WHERE salary_item_id = ? AND employee_id IN ({placeholders})", [salary_item_id] + employee_ids)
        data_tuples = [(emp_id, salary_item_id, amount, start_date, end_date, note) for emp_id in employee_ids]
        cursor.executemany("INSERT INTO employee_salary_item (employee_id, salary_item_id, amount, start_date, end_date, note) VALUES (?, ?, ?, ?, ?, ?)", data_tuples)
        conn.commit()
        return len(data_tuples)
    except Exception as e:
        conn.rollback(); raise e

def batch_update_employee_salary_items(conn, employee_ids, salary_item_id, new_data):
    cursor = conn.cursor()
    try:
        start_date_str = new_data['start_date'].strftime('%Y-%m-%d') if new_data['start_date'] else None
        end_date_str = new_data['end_date'].strftime('%Y-%m-%d') if new_data['end_date'] else None
        cursor.execute("BEGIN TRANSACTION")
        placeholders = ','.join('?' for _ in employee_ids)
        sql_update = f"UPDATE employee_salary_item SET amount = ?, start_date = ?, end_date = ?, note = ? WHERE salary_item_id = ? AND employee_id IN ({placeholders})"
        params = [new_data['amount'], start_date_str, end_date_str, new_data['note'], salary_item_id] + employee_ids
        cursor.execute(sql_update, params)
        conn.commit()
        return cursor.rowcount
    except Exception as e:
        conn.rollback(); raise e

def update_employee_salary_item(conn, record_id, data):
    cursor = conn.cursor()
    sql = "UPDATE employee_salary_item SET amount = ?, start_date = ?, end_date = ?, note = ? WHERE id = ?"
    cursor.execute(sql, (data['amount'], data['start_date'], data['end_date'], data['note'], record_id))
    conn.commit()
    return cursor.rowcount

def delete_employee_salary_item(conn, record_id):
    cursor = conn.cursor()
    sql = "DELETE FROM employee_salary_item WHERE id = ?"
    cursor.execute(sql, (record_id,))
    conn.commit()
    return cursor.rowcount

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
