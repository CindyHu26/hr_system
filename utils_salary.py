# utils_salary.py
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


# --- 薪資項目相關函式 (Salary Item) ---

def get_all_salary_items(conn, active_only=False):
    """取得所有薪資項目"""
    query = "SELECT * FROM salary_item"
    if active_only:
        query += " WHERE is_active = 1"
    query += " ORDER BY type, id"
    return pd.read_sql_query(query, conn)

def add_salary_item(conn, data):
    """新增薪資項目"""
    cursor = conn.cursor()
    sql = "INSERT INTO salary_item (name, type, is_active) VALUES (?, ?, ?)"
    cursor.execute(sql, (data['name'], data['type'], data['is_active']))
    conn.commit()

def update_salary_item(conn, item_id, data):
    """更新薪資項目"""
    cursor = conn.cursor()
    sql = "UPDATE salary_item SET name = ?, type = ?, is_active = ? WHERE id = ?"
    cursor.execute(sql, (data['name'], data['type'], data['is_active'], item_id))
    conn.commit()

def delete_salary_item(conn, item_id):
    """刪除指定的薪資項目"""
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON;")
    sql = "DELETE FROM salary_item WHERE id = ?"
    cursor.execute(sql, (item_id,))
    conn.commit()
    return cursor.rowcount


# --- 員工底薪/眷屬異動歷史相關函式 (Salary Base History) ---

def get_salary_base_history(conn):
    """取得所有員工的底薪/眷屬異動歷史"""
    query = """
    SELECT
        sh.id, sh.employee_id, e.name_ch, sh.base_salary,
        sh.dependents, sh.start_date, sh.end_date, sh.note
    FROM salary_base_history sh
    JOIN employee e ON sh.employee_id = e.id
    ORDER BY e.name_ch, sh.start_date DESC
    """
    return pd.read_sql_query(query, conn)

def add_salary_base_history(conn, data):
    """新增一筆底薪/眷屬異動歷史"""
    cursor = conn.cursor()
    sql = """
    INSERT INTO salary_base_history
    (employee_id, base_salary, dependents, start_date, end_date, note)
    VALUES (?, ?, ?, ?, ?, ?)
    """
    start_date_str = data['start_date'].strftime('%Y-%m-%d') if data['start_date'] else None
    end_date_str = data['end_date'].strftime('%Y-%m-%d') if data['end_date'] else None
    
    cursor.execute(sql, (
        data['employee_id'], data['base_salary'], data['dependents'],
        start_date_str, end_date_str, data['note']
    ))
    conn.commit()

def update_salary_base_history(conn, record_id, data):
    """更新指定的底薪/眷屬異動歷史"""
    cursor = conn.cursor()
    sql = """
    UPDATE salary_base_history SET
    base_salary = ?, dependents = ?, start_date = ?, end_date = ?, note = ?
    WHERE id = ?
    """
    start_date_str = data['start_date'].strftime('%Y-%m-%d') if data['start_date'] else None
    end_date_str = data['end_date'].strftime('%Y-%m-%d') if data['end_date'] else None

    cursor.execute(sql, (
        data['base_salary'], data['dependents'], start_date_str,
        end_date_str, data['note'], record_id
    ))
    conn.commit()

def delete_salary_base_history(conn, record_id):
    """刪除指定的底薪/眷屬異動歷史"""
    cursor = conn.cursor()
    sql = "DELETE FROM salary_base_history WHERE id = ?"
    cursor.execute(sql, (record_id,))
    conn.commit()


# --- 勞健保級距相關函式 (Insurance Grade) ---
def parse_labor_insurance_excel(file_obj):
    """
    (V11 - 加入去重機制) 根據使用者提供的精準行號和欄位邏輯，解析官方 Excel 檔案
    """
    try:
        df = pd.read_excel(file_obj, header=None, engine='xlrd')

        # 定位資料所在的精準行號 (iloc 是 0-based)
        grade_row_data = df.iloc[36]
        salary_row_data = df.iloc[37]
        fee_row_data = df.iloc[68]
        
        # 定位「全時勞工」的起始「欄」，關鍵是找到 "第1級"
        start_col_index = -1
        for i, grade_text in enumerate(grade_row_data):
            if isinstance(grade_text, str) and "第1級" in grade_text:
                start_col_index = i
                break
        if start_col_index == -1:
            raise ValueError("在第37列中找不到 '第1級'，無法定位全時勞工級距表。")

        # 組合資料
        records = []
        for i in range(start_col_index, len(salary_row_data)):
            salary = salary_row_data.get(i)
            grade_text = grade_row_data.get(i)
            
            if pd.notna(salary) and isinstance(salary, (int, float)) and isinstance(grade_text, str):
                try:
                    grade = int(''.join(filter(str.isdigit, grade_text)))
                    employee_fee = fee_row_data.get(i)
                    employer_fee = fee_row_data.get(i + 1)
                    records.append({
                        'grade': grade,
                        'salary_max': salary,
                        'employee_fee': employee_fee,
                        'employer_fee': employer_fee,
                    })
                except (ValueError, TypeError):
                    continue # 如果無法轉換為數字，則跳過此欄

        if not records:
            raise ValueError("無法從指定的行號中提取有效的級距資料。請確認檔案格式未變。")

        df_final = pd.DataFrame(records)
        df_final.dropna(subset=['grade', 'salary_max'], inplace=True)
        
        # **核心修正：加入去重保險機制**
        # 根據 'grade' 欄位去除重複項，保留第一個出現的
        df_final.drop_duplicates(subset=['grade'], keep='first', inplace=True)
        
        # 計算 salary_min
        df_final['salary_min'] = df_final['salary_max'].shift(1).fillna(0) + 1
        df_final.loc[df_final.index[0], 'salary_min'] = 0
        
        return df_final[['grade', 'salary_min', 'salary_max', 'employee_fee', 'employer_fee']]

    except Exception as e:
        raise ValueError(f"解析勞保 Excel 檔案時發生錯誤: {e}")


def parse_insurance_html_table(html_content):
    """(V5版 - 加入去重機制) 從 HTML 文本中解析健保級距表"""
    try:
        tables = pd.read_html(io.StringIO(html_content))
        target_df = next((df for df in tables if '月投保金額' in ''.join(map(str, df.columns))), None)
        if target_df is None:
            raise ValueError("在 HTML 中找不到包含 '月投保金額' 的表格。")

        df = target_df.copy()
        df.columns = ['_'.join(map(str, col)).strip() for col in df.columns.values]
        df.ffill(inplace=True)

        rename_map = {
            df.columns[0]: 'grade',
            df.columns[1]: 'salary_max',
            df.columns[2]: 'employee_fee',
            df.columns[6]: 'employer_fee',
            df.columns[7]: 'gov_fee'
        }
        df.rename(columns=rename_map, inplace=True)
        
        cols_to_numeric = ['grade', 'salary_max', 'employee_fee', 'employer_fee', 'gov_fee']
        for col in cols_to_numeric:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(r'[\s,元$]', '', regex=True)
                df[col] = pd.to_numeric(df[col], errors='coerce')

        df.dropna(subset=['grade', 'salary_max'], inplace=True)
        
        # **核心修正：加入去重保險機制**
        df.drop_duplicates(subset=['grade'], keep='first', inplace=True)
        
        df['salary_min'] = df['salary_max'].shift(1).fillna(0) + 1
        df.loc[df.index[0], 'salary_min'] = 0
        
        final_cols = ['grade', 'salary_min', 'salary_max', 'employee_fee', 'employer_fee', 'gov_fee']
        return df[[col for col in final_cols if col in df.columns]]

    except Exception as e:
        raise ValueError(f"解析 HTML 時發生錯誤: {e}")

def get_insurance_grades(conn):
    return pd.read_sql_query("SELECT * FROM insurance_grade ORDER BY start_date DESC, type, grade", conn)

def batch_insert_insurance_grades(conn, df, grade_type, start_date):
    cursor = conn.cursor()
    try:
        start_date_str = start_date.strftime('%Y-%m-%d')
        cursor.execute("BEGIN TRANSACTION")
        cursor.execute("DELETE FROM insurance_grade WHERE type = ? AND start_date = ?", (grade_type, start_date_str))
        required_cols = {
            'grade': 0, 'salary_min': 0, 'salary_max': 0,
            'employee_fee': None, 'employer_fee': None, 'gov_fee': None, 'note': None
        }
        for col, default in required_cols.items():
            if col not in df.columns:
                df[col] = default
        df_to_insert = df[list(required_cols.keys())].copy()
        df_to_insert['type'] = grade_type
        df_to_insert['start_date'] = start_date_str
        cols_for_sql = ['start_date', 'type'] + list(required_cols.keys())
        placeholders = ','.join(['?'] * len(cols_for_sql))
        sql = f"INSERT INTO insurance_grade ({','.join(cols_for_sql)}) VALUES ({placeholders})"
        data_tuples = [tuple(row) for row in df_to_insert[cols_for_sql].to_numpy()]
        cursor.executemany(sql, data_tuples)
        conn.commit()
        return len(data_tuples)
    except Exception as e:
        conn.rollback()
        raise e

def update_insurance_grade(conn, record_id, data):
    """更新單筆級距資料"""
    cursor = conn.cursor()
    sql = """
    UPDATE insurance_grade SET
    salary_min = ?, salary_max = ?, employee_fee = ?,
    employer_fee = ?, gov_fee = ?, note = ?
    WHERE id = ?
    """
    cursor.execute(sql, (
        data['salary_min'], data['salary_max'], data['employee_fee'],
        data['employer_fee'], data['gov_fee'], data['note'],
        record_id
    ))
    conn.commit()

def delete_insurance_grade(conn, record_id):
    """刪除單筆級距資料"""
    cursor = conn.cursor()
    sql = "DELETE FROM insurance_grade WHERE id = ?"
    cursor.execute(sql, (record_id,))
    conn.commit()

# --- 員工常態薪資項設定相關函式 (Employee Salary Item) ---

def get_employee_salary_items(conn):
    """取得所有員工的常態薪資項設定"""
    query = """
    SELECT
        esi.id,
        e.id as employee_id,
        e.name_ch as '員工姓名',
        si.id as salary_item_id,
        si.name as '項目名稱',
        si.type as '類型',
        esi.amount as '金額',
        esi.start_date as '生效日',
        esi.end_date as '結束日',
        esi.note as '備註'
    FROM employee_salary_item esi
    JOIN employee e ON esi.employee_id = e.id
    JOIN salary_item si ON esi.salary_item_id = si.id
    ORDER BY e.name_ch, si.name
    """
    return pd.read_sql_query(query, conn)

def get_settings_grouped_by_amount(conn, salary_item_id):
    """
    根據薪資項目ID，查詢所有相關設定，並按金額分組。
    返回一個字典，鍵是金額，值是擁有該金額設定的員工列表。
    """
    if not salary_item_id:
        return {}
    
    query = """
    SELECT
        esi.amount,
        e.id as employee_id,
        e.name_ch
    FROM employee_salary_item esi
    JOIN employee e ON esi.employee_id = e.id
    WHERE esi.salary_item_id = ?
    ORDER BY esi.amount, e.name_ch
    """
    df = pd.read_sql_query(query, conn, params=(int(salary_item_id),))
    
    # 按金額分組
    grouped_data = {}
    if not df.empty:
        for amount, group in df.groupby('amount'):
            grouped_data[amount] = group[['employee_id', 'name_ch']].to_dict('records')
            
    return grouped_data


def batch_add_employee_salary_items(conn, employee_ids, salary_item_id, amount, start_date, end_date, note):
    """
    批次為多位員工新增一筆常態薪資項。
    此操作會先檢查並刪除這些員工在同一項目上的舊設定，再插入新設定。
    """
    cursor = conn.cursor()
    try:
        start_date_str = start_date.strftime('%Y-%m-%d')
        end_date_str = end_date.strftime('%Y-%m-%d') if end_date else None
        
        cursor.execute("BEGIN TRANSACTION")
        
        # 為了簡化，我們先刪除這些員工在此項目上的舊設定
        placeholders = ','.join('?' for _ in employee_ids)
        sql_delete = f"DELETE FROM employee_salary_item WHERE salary_item_id = ? AND employee_id IN ({placeholders})"
        cursor.execute(sql_delete, [salary_item_id] + employee_ids)

        # 批次插入新設定
        sql_insert = """
        INSERT INTO employee_salary_item
        (employee_id, salary_item_id, amount, start_date, end_date, note)
        VALUES (?, ?, ?, ?, ?, ?)
        """
        data_tuples = [
            (emp_id, salary_item_id, amount, start_date_str, end_date_str, note)
            for emp_id in employee_ids
        ]
        cursor.executemany(sql_insert, data_tuples)
        
        conn.commit()
        return len(data_tuples)
        
    except Exception as e:
        conn.rollback()
        raise e
    
def batch_update_employee_salary_items(conn, employee_ids, salary_item_id, new_data):
    """
    批次為指定的多位員工，更新某個常態薪資項的設定。
    """
    cursor = conn.cursor()
    try:
        start_date_str = new_data['start_date'].strftime('%Y-%m-%d') if new_data['start_date'] else None
        end_date_str = new_data['end_date'].strftime('%Y-%m-%d') if new_data['end_date'] else None
        
        cursor.execute("BEGIN TRANSACTION")
        
        placeholders = ','.join('?' for _ in employee_ids)
        sql_update = f"""
        UPDATE employee_salary_item SET
        amount = ?, start_date = ?, end_date = ?, note = ?
        WHERE salary_item_id = ? AND employee_id IN ({placeholders})
        """
        
        # 準備參數
        params = [new_data['amount'], start_date_str, end_date_str, new_data['note'], salary_item_id] + employee_ids
        
        cursor.execute(sql_update, params)
        
        conn.commit()
        # 回傳受影響的行數
        return cursor.rowcount

    except Exception as e:
        conn.rollback()
        raise e

def update_employee_salary_item(conn, record_id, data):
    """更新指定的單筆常態薪資項設定"""
    cursor = conn.cursor()
    sql = """
    UPDATE employee_salary_item SET
    amount = ?, start_date = ?, end_date = ?, note = ?
    WHERE id = ?
    """
    start_date_str = data['start_date'].strftime('%Y-%m-%d') if data['start_date'] else None
    end_date_str = data['end_date'].strftime('%Y-%m-%d') if data['end_date'] else None

    cursor.execute(sql, (
        data['amount'], start_date_str, end_date_str, data['note'],
        record_id
    ))
    conn.commit()
    return cursor.rowcount

def delete_employee_salary_item(conn, record_id):
    """刪除指定的單筆常態薪資項設定"""
    cursor = conn.cursor()
    sql = "DELETE FROM employee_salary_item WHERE id = ?"
    cursor.execute(sql, (record_id,))
    conn.commit()
    return cursor.rowcount

# --- 薪資單產生與計算核心函式 (Salary Calculation) ---

def check_salary_records_exist(conn, year, month):
    """檢查指定年月的薪資主紀錄是否存在"""
    query = "SELECT 1 FROM salary WHERE year = ? AND month = ? LIMIT 1"
    return conn.cursor().execute(query, (year, month)).fetchone() is not None

def get_active_employees_for_month(conn, year, month):
    """取得在指定年月仍在職的員工"""
    month_first_day = f"{year}-{month:02d}-01"
    _, last_day = pd.Timestamp(year, month, 1).days_in_month
    month_last_day = f"{year}-{month:02d}-{last_day}"

    query = """
    SELECT id, name_ch FROM employee
    WHERE (entry_date IS NOT NULL AND entry_date <= ?)
      AND (resign_date IS NULL OR resign_date >= ?)
    """
    return pd.read_sql_query(query, conn, params=(month_last_day, month_first_day))

def generate_initial_salary_records(conn, year, month):
    """為所有在職員工產生薪資單主紀錄和所有明細的初始計算"""
    employees_df = get_active_employees_for_month(conn, year, month)
    if employees_df.empty:
        return 0
    
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN TRANSACTION")
        
        for _, emp in employees_df.iterrows():
            # 1. 建立薪資主紀錄 (salary)
            sql_salary = "INSERT INTO salary (employee_id, year, month) VALUES (?, ?, ?)"
            cursor.execute(sql_salary, (emp['id'], year, month))
            salary_id = cursor.lastrowid

            # 2. 計算並插入薪資明細 (salary_detail)
            
            # a. 取得底薪與眷屬數 (從 salary_base_history)
            sql_base = "SELECT base_salary, dependents FROM salary_base_history WHERE employee_id = ? AND start_date <= date('now') ORDER BY start_date DESC LIMIT 1"
            base_info = cursor.execute(sql_base, (emp['id'],)).fetchone()
            base_salary = base_info[0] if base_info else 0
            dependents = base_info[1] if base_info else 0

            # b. 取得所有常態薪資項 (從 employee_salary_item)
            sql_items = """
            SELECT salary_item_id, amount FROM employee_salary_item
            WHERE employee_id = ? AND start_date <= date('now')
              AND (end_date IS NULL OR end_date >= date('now'))
            """
            recurring_items = cursor.execute(sql_items, (emp['id'],)).fetchall()
            
            # c. 取得勞健保級距與費用
            salary_for_insurance = base_salary # 簡化：先用底薪判斷級距
            
            # 勞保
            sql_labor = "SELECT employee_fee FROM insurance_grade WHERE type = 'labor' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            labor_fee_tuple = cursor.execute(sql_labor, (salary_for_insurance,)).fetchone()
            labor_fee = labor_fee_tuple[0] if labor_fee_tuple else 0
            
            # 健保 (需要考慮眷屬)
            sql_health = "SELECT employee_fee FROM insurance_grade WHERE type = 'health' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            health_fee_per_person = cursor.execute(sql_health, (salary_for_insurance,)).fetchone()
            
            # 健保費 = 個人保費 * (1 + 眷屬數(上限3))
            num_insured = 1 + min(dependents, 3)
            health_fee = (health_fee_per_person[0] * num_insured) if health_fee_per_person else 0

            # d. 準備所有要插入的薪資明細
            details_to_insert = []
            # - 底薪 (假設 "底薪" 的 salary_item_id 為 1)
            details_to_insert.append((salary_id, 1, base_salary))
            # - 常態項目
            for item_id, amount in recurring_items:
                details_to_insert.append((salary_id, item_id, amount))
            # - 勞保費 (假設 "勞保費" 的 salary_item_id 為 2)
            details_to_insert.append((salary_id, 2, -labor_fee)) # 扣款為負數
            # - 健保費 (假設 "健保費" 的 salary_item_id 為 3)
            details_to_insert.append((salary_id, 3, -health_fee)) # 扣款為負數

            # 3. 批次插入明細
            sql_detail_insert = "INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)"
            cursor.executemany(sql_detail_insert, details_to_insert)

        conn.commit()
        return len(employees_df)
    except Exception as e:
        conn.rollback()
        raise e

def get_salary_report_for_editing(conn, year, month):
    """取得指定年月的薪資報表，格式為 員工 vs 薪資項目，適合 data_editor"""
    query = """
    SELECT
        s.id as salary_id,
        e.name_ch as "員工姓名",
        si.name as item_name,
        sd.amount
    FROM salary s
    JOIN employee e ON s.employee_id = e.id
    LEFT JOIN salary_detail sd ON s.id = sd.salary_id
    LEFT JOIN salary_item si ON sd.salary_item_id = si.id
    WHERE s.year = ? AND s.month = ?
    ORDER BY e.name_ch, si.id
    """
    df = pd.read_sql_query(query, conn, params=(year, month))
    if df.empty:
        return pd.DataFrame()
    
    # 使用 pivot_table 將資料轉換為寬表格
    pivot_df = df.pivot_table(index=['salary_id', '員工姓名'], columns='item_name', values='amount').reset_index()
    return pivot_df

def update_salary_detail_by_name(conn, salary_id, item_name, new_amount):
    """根據 salary_id 和項目名稱來更新或插入金額"""
    cursor = conn.cursor()
    # 1. 取得 salary_item_id
    item_id_tuple = cursor.execute("SELECT id FROM salary_item WHERE name = ?", (item_name,)).fetchone()
    if not item_id_tuple:
        raise ValueError(f"找不到名為 '{item_name}' 的薪資項目。")
    item_id = item_id_tuple[0]

    # 2. 檢查此明細是否存在
    detail_id_tuple = cursor.execute("SELECT id FROM salary_detail WHERE salary_id = ? AND salary_item_id = ?", (salary_id, item_id)).fetchone()
    
    if detail_id_tuple:
        # 更新
        sql = "UPDATE salary_detail SET amount = ? WHERE id = ?"
        cursor.execute(sql, (new_amount, detail_id_tuple[0]))
    else:
        # 新增
        sql = "INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)"
        cursor.execute(sql, (salary_id, item_id, new_amount))
        
    conn.commit()
    return cursor.rowcount

def check_salary_records_exist(conn, year, month):
    """檢查指定年月的薪資主紀錄是否存在"""
    query = "SELECT 1 FROM salary WHERE year = ? AND month = ? LIMIT 1"
    return conn.cursor().execute(query, (year, month)).fetchone() is not None

def get_active_employees_for_month(conn, year, month):
    """取得在指定年月仍在職的員工"""
    month_first_day = f"{year}-{month:02d}-01"
    _, last_day = pd.Timestamp(year, month, 1).days_in_month
    month_last_day = f"{year}-{month:02d}-{last_day}"

    query = """
    SELECT id, name_ch FROM employee
    WHERE (entry_date IS NOT NULL AND entry_date <= ?)
      AND (resign_date IS NULL OR resign_date >= ?)
    """
    return pd.read_sql_query(query, conn, params=(month_last_day, month_first_day))

def generate_initial_salary_records(conn, year, month):
    """
    (V3 - 動態ID終極版) 為所有在職員工產生薪資單，並自動計算出勤與請假相關項目
    """
    employees_df = get_active_employees_for_month(conn, year, month)
    if employees_df.empty:
        return 0
    
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN TRANSACTION")
        
        # 1. 讀取薪資規則
        rules_df = pd.read_sql_query("SELECT rule_key, value FROM payroll_rule", conn, index_col='rule_key')
        rules = rules_df['value'].to_dict()
        hourly_rate_divisor = float(rules.get('hourly_rate_divisor', 240))

        # **核心修正：從資料庫動態讀取項目 ID**
        items_df = pd.read_sql_query("SELECT id, name FROM salary_item", conn)
        ITEM_IDS = pd.Series(items_df.id.values, index=items_df.name).to_dict()

        # 檢查核心項目是否存在
        required_items = ['底薪', '勞健保', '加班費', '遲到', '早退', '事假', '病假']
        missing_items = [item for item in required_items if item not in ITEM_IDS]
        if missing_items:
            raise ValueError(f"系統缺少必要的薪資項目，請至「薪資項目管理」頁面新增：{', '.join(missing_items)}")
        
        for _, emp in employees_df.iterrows():
            # 2. 建立薪資主紀錄 (salary)
            sql_salary = "INSERT OR IGNORE INTO salary (employee_id, year, month) VALUES (?, ?, ?)"
            cursor.execute(sql_salary, (emp['id'], year, month))
            salary_id_tuple = cursor.execute("SELECT id FROM salary WHERE employee_id = ? AND year = ? AND month = ?", (emp['id'], year, month)).fetchone()
            if not salary_id_tuple: continue
            salary_id = salary_id_tuple[0]

            # 3. 計算基礎薪資與時薪
            sql_base = "SELECT base_salary, dependents FROM salary_base_history WHERE employee_id = ? AND start_date <= date('now') ORDER BY start_date DESC LIMIT 1"
            base_info = cursor.execute(sql_base, (emp['id'],)).fetchone()
            base_salary = base_info[0] if base_info else 0
            dependents = base_info[1] if base_info else 0
            hourly_rate = base_salary / hourly_rate_divisor if hourly_rate_divisor > 0 else 0

            details_to_insert = []
            # - 底薪
            details_to_insert.append((salary_id, ITEM_IDS['base_salary'], base_salary))

            # 4. 計算常態薪資項
            sql_items = "SELECT salary_item_id, amount FROM employee_salary_item WHERE employee_id = ? AND start_date <= date('now') AND (end_date IS NULL OR end_date >= date('now'))"
            recurring_items = cursor.execute(sql_items, (emp['id'],)).fetchall()
            for item_id, amount in recurring_items:
                details_to_insert.append((salary_id, item_id, amount))

            # **核心修正：合併計算勞健保費**
            salary_for_insurance = base_salary
            sql_labor = "SELECT employee_fee FROM insurance_grade WHERE type = 'labor' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            labor_fee = (cursor.execute(sql_labor, (salary_for_insurance,)).fetchone() or [0])[0]
            
            sql_health = "SELECT employee_fee FROM insurance_grade WHERE type = 'health' AND ? BETWEEN salary_min AND salary_max ORDER BY start_date DESC LIMIT 1"
            health_fee_per_person = (cursor.execute(sql_health, (salary_for_insurance,)).fetchone() or [0])[0]
            num_insured = 1 + min(dependents, 3)
            health_fee = health_fee_per_person * num_insured
            
            total_insurance_fee = labor_fee + health_fee
            details_to_insert.append((salary_id, ITEM_IDS['勞健保'], -total_insurance_fee)) # 扣款為負數

            # **核心修正：更穩健地處理出勤分鐘數，避免 TypeError**
            month_str = f"{year}-{month:02d}"
            # 加班費
            sql_overtime = "SELECT SUM(overtime1_minutes), SUM(overtime2_minutes), SUM(overtime3_minutes) FROM attendance WHERE employee_id = ? AND strftime('%Y-%m', date) = ?"
            overtime_tuple = cursor.execute(sql_overtime, (emp['id'], month_str)).fetchone()
            total_overtime_pay = 0
            if overtime_tuple:
                ot1_mins, ot2_mins, ot3_mins = (overtime_tuple[0] or 0, overtime_tuple[1] or 0, overtime_tuple[2] or 0)
                total_overtime_pay += (ot1_mins / 60) * hourly_rate * float(rules.get('weekday_overtime_rate', 1.34))
                total_overtime_pay += ((ot2_mins + ot3_mins) / 60) * hourly_rate * 1.67
            details_to_insert.append((salary_id, ITEM_IDS['加班費'], round(total_overtime_pay)))
            
            # 遲到/早退
            sql_lateness = "SELECT SUM(late_minutes), SUM(early_leave_minutes) FROM attendance WHERE employee_id = ? AND strftime('%Y-%m', date) = ?"
            lateness_tuple = cursor.execute(sql_lateness, (emp['id'], month_str)).fetchone()
            total_late_deduction, total_early_leave_deduction = 0, 0
            if lateness_tuple:
                late_mins, early_mins = (lateness_tuple[0] or 0, lateness_tuple[1] or 0)
                total_late_deduction = late_mins * (hourly_rate / 60)
                total_early_leave_deduction = early_mins * (hourly_rate / 60)
            details_to_insert.append((salary_id, ITEM_IDS['遲到'], -round(total_late_deduction)))
            details_to_insert.append((salary_id, ITEM_IDS['早退'], -round(total_early_leave_deduction)))
            
            # 請假
            sql_leave = "SELECT leave_type, SUM(duration) FROM leave_record WHERE employee_id = ? AND strftime('%Y-%m', start_date) = ? GROUP BY leave_type"
            leave_hours = cursor.execute(sql_leave, (emp['id'], month_str)).fetchall()
            total_personal_leave_deduction, total_sick_leave_deduction = 0, 0
            if leave_hours:
                for leave_type, hours in leave_hours:
                    if leave_type == '事假':
                        total_personal_leave_deduction += hours * hourly_rate
                    elif leave_type == '病假':
                        total_sick_leave_deduction += (hours * hourly_rate * 0.5)
            details_to_insert.append((salary_id, ITEM_IDS['事假'], -round(total_personal_leave_deduction)))
            details_to_insert.append((salary_id, ITEM_IDS['病假'], -round(total_sick_leave_deduction)))

            # 7. 批次插入所有明細
            sql_detail_insert = "INSERT OR IGNORE INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)"
            cursor.executemany(sql_detail_insert, details_to_insert)

        conn.commit()
        return len(employees_df)
    except Exception as e:
        conn.rollback()
        raise e