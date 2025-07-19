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
    query = "SELECT * FROM salary_item ORDER BY type, id"
    if active_only:
        query = "SELECT * FROM salary_item WHERE is_active = 1 ORDER BY type, id"
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
    query = """
    SELECT sh.id, sh.employee_id, e.name_ch, sh.base_salary,
           sh.dependents, sh.start_date, sh.end_date, sh.note
    FROM salary_base_history sh JOIN employee e ON sh.employee_id = e.id
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

# ---【新函式 1】篩選底薪低於新標準的員工 ---
def get_employees_below_minimum_wage(conn, new_minimum_wage: int):
    """
    查詢所有在職且目前底薪低於新基本工資的員工。
    """
    query = """
    WITH latest_salary AS (
        SELECT
            employee_id,
            base_salary,
            dependents,
            -- 使用 ROW_NUMBER() 找出每位員工最新的薪資紀錄
            ROW_NUMBER() OVER(PARTITION BY employee_id ORDER BY start_date DESC) as rn
        FROM salary_base_history
    )
    SELECT
        e.id as employee_id,
        e.name_ch as "員工姓名",
        ls.base_salary as "目前底薪",
        ls.dependents as "目前眷屬數"
    FROM employee e
    LEFT JOIN latest_salary ls ON e.id = ls.employee_id
    WHERE
        e.resign_date IS NULL -- 只找出在職員工
        AND ls.rn = 1 -- 確保只用最新的薪資紀錄來比較
        AND ls.base_salary < ? -- 核心條件：底薪低於新標準
    ORDER BY e.id;
    """
    return pd.read_sql_query(query, conn, params=(new_minimum_wage,))

# ---【新函式 2】批次更新員工底薪 ---
def batch_update_basic_salary(conn, preview_df: pd.DataFrame, new_wage: int, effective_date):
    """
    為預覽列表中的所有員工，批次新增一筆調薪紀錄。
    """
    cursor = conn.cursor()
    try:
        data_to_insert = []
        for _, row in preview_df.iterrows():
            data_to_insert.append((
                row['employee_id'],
                new_wage, # 新的底薪
                row['目前眷屬數'], # 眷屬數維持不變
                effective_date.strftime('%Y-%m-%d'),
                None, # 結束日為空
                f"配合 {effective_date.year} 年基本工資調整" # 自動產生備註
            ))
        
        sql = """
        INSERT INTO salary_base_history
        (employee_id, base_salary, dependents, start_date, end_date, note)
        VALUES (?, ?, ?, ?, ?, ?)
        """
        cursor.executemany(sql, data_to_insert)
        conn.commit()
        return cursor.rowcount
    except Exception as e:
        conn.rollback()
        raise e

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
    query = "SELECT 1 FROM salary WHERE year = ? AND month = ? LIMIT 1"
    return conn.cursor().execute(query, (year, month)).fetchone() is not None

def get_active_employees_for_month(conn, year, month):
    month_first_day = f"{year}-{month:02d}-01"
    last_day = pd.Timestamp(year, month, 1).days_in_month
    month_last_day = f"{year}-{month:02d}-{last_day}"
    query = """
    SELECT id, name_ch FROM employee
    WHERE (entry_date IS NOT NULL AND entry_date <= ?)
      AND (resign_date IS NULL OR resign_date >= ?)
    ORDER BY id ASC
    """
    return pd.read_sql_query(query, conn, params=(month_last_day, month_first_day))

def update_salary_detail_by_name(conn, salary_id, item_name, new_amount):
    cursor = conn.cursor()
    item_id_tuple = cursor.execute("SELECT id FROM salary_item WHERE name = ?", (item_name,)).fetchone()
    if not item_id_tuple:
        raise ValueError(f"Could not find salary item named '{item_name}'.")
    item_id = item_id_tuple[0]
    detail_id_tuple = cursor.execute("SELECT id FROM salary_detail WHERE salary_id = ? AND salary_item_id = ?", (salary_id, item_id)).fetchone()
    if detail_id_tuple:
        sql = "UPDATE salary_detail SET amount = ? WHERE id = ?"
        cursor.execute(sql, (new_amount, detail_id_tuple[0]))
    else:
        sql = "INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)"
        cursor.execute(sql, (salary_id, item_id, new_amount))
    conn.commit()
    return cursor.rowcount

# ---【核心修正】新增一個函式，專門用來取得薪資項目類型 ---
def get_item_types(conn):
    """獲取所有薪資項目的名稱及其類型 (earning/deduction) 的字典"""
    return pd.read_sql("SELECT name, type FROM salary_item", conn).set_index('name')['type'].to_dict()

# --- 薪資試算引擎現在會回傳計算結果和項目類型 ---
def calculate_salary_df(conn, year, month, non_insured_names: list = None):
    """V10 薪資試算引擎: 新增三大計算欄位，並回傳項目類型以供上色"""
    # ... (此函式前半部分的計算邏輯保持不變) ...
    if non_insured_names is None: non_insured_names = []
    employees_df = get_active_employees_for_month(conn, year, month)
    if employees_df.empty: return pd.DataFrame(), {}
    
    item_types = get_item_types(conn)
    all_salary_data = []

    for _, emp in employees_df.iterrows():
        # ... (員工薪資明細的計算邏輯完全不變) ...
        # ... (包括底薪、勞健保、常態津貼、事病假等) ...
        pass # 此處省略重複程式碼

    if not all_salary_data: return pd.DataFrame(), {}

    result_df = pd.DataFrame(all_salary_data).fillna(0)
    for col in result_df.columns:
        if col != '員工姓名':
            result_df[col] = result_df[col].round().astype(int)

    # 計算總計
    earning_cols = [c for c, t in item_types.items() if t == 'earning' and c in result_df.columns]
    deduction_cols = [c for c, t in item_types.items() if t == 'deduction' and c in result_df.columns]
    
    result_df['應發總額'] = result_df[earning_cols].sum(axis=1)
    result_df['應扣總額'] = result_df[deduction_cols].sum(axis=1)
    result_df['實發淨薪'] = result_df['應發總額'] + result_df['應扣總額']
    
    # ---【核心修正】新增三大計算欄位 ---
    result_df['申報薪資'] = (
        result_df.get('底薪', 0) + 
        result_df.get('事假', 0) + result_df.get('病假', 0) + 
        result_df.get('遲到', 0) + result_df.get('早退', 0)
    )
    result_df['匯入銀行'] = (
        result_df.get('底薪', 0) + result_df.get('加班費', 0) +
        result_df.get('勞健保', 0) + 
        result_df.get('事假', 0) + result_df.get('病假', 0) +
        result_df.get('遲到', 0) + result_df.get('早退', 0)
    )
    result_df['現金'] = result_df['實發淨薪'] - result_df['匯入銀行']
    
    # 調整欄位順序
    final_cols = ['員工姓名', '實發淨薪', '應發總額', '應扣總額', '申報薪資', '匯入銀行', '現金'] + earning_cols + deduction_cols
    
    # 回傳 DataFrame 和 項目類型字典
    return result_df[[c for c in final_cols if c in result_df.columns]], item_types

def save_salary_df(conn, year, month, df: pd.DataFrame):
    """將試算的 DataFrame 寫入資料庫"""
    cursor = conn.cursor()
    emp_map = pd.read_sql("SELECT id, name_ch FROM employee", conn).set_index('name_ch')['id'].to_dict()
    item_map = pd.read_sql("SELECT id, name FROM salary_item", conn).set_index('name')['id'].to_dict()

    for _, row in df.iterrows():
        emp_name = row['員工姓名']
        emp_id = emp_map.get(emp_name)
        if not emp_id: continue

        cursor.execute("INSERT OR IGNORE INTO salary (employee_id, year, month) VALUES (?, ?, ?)", (emp_id, year, month))
        salary_id = cursor.execute("SELECT id FROM salary WHERE employee_id = ? AND year = ? AND month = ?", (emp_id, year, month)).fetchone()[0]
        
        cursor.execute("DELETE FROM salary_detail WHERE salary_id = ?", (salary_id,))

        details_to_insert = []
        for item_name, amount in row.items():
            if item_name in ['員工姓名', '實發淨薪', '應發總額', '應扣總額']: continue
            item_id = item_map.get(item_name)
            if item_id and amount != 0:
                details_to_insert.append((salary_id, item_id, int(amount)))
        
        if details_to_insert:
            sql = "INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)"
            cursor.executemany(sql, details_to_insert)
    conn.commit()

def get_previous_non_insured_names(conn, current_year, current_month) -> list:
    """從資料庫中，取得指定月份的上一個月中，勞健保為0的員工名單"""
    cursor = conn.cursor()
    # 計算上一個月的年和月
    prev_date = (datetime(current_year, current_month, 1) - pd.DateOffset(months=1))
    year, month = prev_date.year, prev_date.month

    item_id_insurance_tuple = cursor.execute("SELECT id FROM salary_item WHERE name = '勞健保'").fetchone()
    if not item_id_insurance_tuple: return []
    item_id_insurance = item_id_insurance_tuple[0]
    
    query = """
    SELECT e.name_ch
    FROM employee e
    WHERE e.id IN (
        SELECT s.employee_id
        FROM salary s
        LEFT JOIN salary_detail sd ON s.id = sd.salary_id AND sd.salary_item_id = ?
        WHERE s.year = ? AND s.month = ?
        GROUP BY s.employee_id
        HAVING IFNULL(SUM(sd.amount), 0) = 0
    )
    """
    names = cursor.execute(query, (item_id_insurance, year, month)).fetchall()
    return [name[0] for name in names]

# ---【核心修正】強化的批次更新函式，自動處理正負值 ---
def batch_update_salary_details_from_excel(conn, year, month, uploaded_file) -> dict:
    """
    V9 從Excel批次更新薪資明細，自動根據項目類型轉換正負值。
    """
    report = {"success": [], "skipped_emp": [], "skipped_item": []}
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        raise ValueError(f"讀取 Excel 檔案失敗: {e}")

    cursor = conn.cursor()
    
    # 建立員工姓名 -> ID 的映射
    emp_map = pd.read_sql("SELECT id, name_ch FROM employee", conn).set_index('name_ch')['id'].to_dict()
    
    # 建立薪資項目名稱 -> (ID, 類型) 的映射
    item_map_df = pd.read_sql("SELECT id, name, type FROM salary_item", conn)
    item_map = {row['name']: {'id': row['id'], 'type': row['type']} for index, row in item_map_df.iterrows()}

    
    for index, row in df.iterrows():
        emp_name = row.get('員工姓名')
        if not emp_name: continue

        emp_id = emp_map.get(emp_name)
        if not emp_id:
            report["skipped_emp"].append(emp_name)
            continue
        
        salary_id_tuple = cursor.execute("SELECT id FROM salary WHERE employee_id = ? AND year = ? AND month = ?", (emp_id, year, month)).fetchone()
        if not salary_id_tuple: continue
        salary_id = salary_id_tuple[0]
        
        for item_name, amount in row.items():
            if item_name == '員工姓名' or pd.isna(amount): continue
            
            item_info = item_map.get(item_name)
            if not item_info:
                report["skipped_item"].append(item_name)
                continue
            
            item_id = item_info['id']
            item_type = item_info['type']
            
            # ---【智慧轉換核心邏輯】---
            # 如果是扣除項，確保金額為負；如果是給付項，確保金額為正。
            final_amount = -abs(float(amount)) if item_type == 'deduction' else abs(float(amount))
            
            detail_id_tuple = cursor.execute("SELECT id FROM salary_detail WHERE salary_id = ? AND salary_item_id = ?", (salary_id, item_id)).fetchone()
            
            if detail_id_tuple:
                cursor.execute("UPDATE salary_detail SET amount = ? WHERE id = ?", (final_amount, detail_id_tuple[0]))
            else:
                cursor.execute("INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)", (salary_id, item_id, final_amount))
            
            report["success"].append(f"{emp_name} - {item_name}")
            
    conn.commit()
    # 去除重複項
    report["skipped_emp"] = list(set(report["skipped_emp"]))
    report["skipped_item"] = list(set(report["skipped_item"]))
    return report

def save_data_editor_changes(conn, year, month, edited_df: pd.DataFrame):
    """
    將 st.data_editor 編輯後的 DataFrame 儲存回資料庫。
    此函式邏輯與批次上傳非常相似。
    """
    cursor = conn.cursor()
    
    emp_map = pd.read_sql("SELECT id, name_ch FROM employee", conn).set_index('name_ch')['id'].to_dict()
    item_map = pd.read_sql("SELECT id, name FROM salary_item", conn).set_index('name')['id'].to_dict()

    for _, row in edited_df.iterrows():
        emp_name = row.get('員工姓名')
        emp_id = emp_map.get(emp_name)
        if not emp_id: continue

        salary_id_tuple = cursor.execute("SELECT id FROM salary WHERE employee_id = ? AND year = ? AND month = ?", (emp_id, year, month)).fetchone()
        if not salary_id_tuple: continue
        salary_id = salary_id_tuple[0]

        for item_name, amount in row.items():
            if item_name == '員工姓名' or pd.isna(amount): continue
            
            item_id = item_map.get(item_name)
            if not item_id: continue

            # 更新對應的 salary_detail 記錄
            cursor.execute(
                "UPDATE salary_detail SET amount = ? WHERE salary_id = ? AND salary_item_id = ?",
                (amount, salary_id, item_id)
            )

    conn.commit()

def get_salary_report_for_editing(conn, year, month):
    """V13 讀取已儲存的薪資單，並確保依照員工ID排序"""
    query = """
    SELECT e.id as employee_id, e.name_ch as "員工姓名", si.name as item_name, sd.amount
    FROM salary s
    JOIN employee e ON s.employee_id = e.id
    JOIN salary_detail sd ON s.id = sd.salary_id
    JOIN salary_item si ON sd.salary_item_id = si.id
    WHERE s.year = ? AND s.month = ?
    """
    df = pd.read_sql_query(query, conn, params=(year, month))
    item_types = get_item_types(conn)

    if df.empty:
        return pd.DataFrame(), item_types

    pivot_df = df.pivot_table(index=['employee_id', '員工姓名'], columns='item_name', values='amount').reset_index().fillna(0)
    pivot_df.sort_values('employee_id', inplace=True)
    pivot_df.drop(columns=['employee_id'], inplace=True) # 排序後即可移除 ID
    
    # 計算總計
    earning_cols = [c for c, t in item_types.items() if t == 'earning' and c in pivot_df.columns]
    deduction_cols = [c for c, t in item_types.items() if t == 'deduction' and c in pivot_df.columns]
    
    pivot_df['應發總額'] = pivot_df[earning_cols].sum(axis=1, numeric_only=True)
    pivot_df['應扣總額'] = pivot_df[deduction_cols].sum(axis=1, numeric_only=True)
    pivot_df['實發淨薪'] = pivot_df['應發總額'] + pivot_df['應扣總額']
    
    # 新增三大計算欄位
    pivot_df['申報薪資'] = (
        pivot_df.get('底薪', 0) + 
        pivot_df.get('事假', 0) + pivot_df.get('病假', 0) + 
        pivot_df.get('遲到', 0) + pivot_df.get('早退', 0)
    ).fillna(0)
    pivot_df['匯入銀行'] = (
        pivot_df.get('底薪', 0) + pivot_df.get('加班費', 0) +
        pivot_df.get('勞健保', 0) + 
        pivot_df.get('事假', 0) + pivot_df.get('病假', 0) +
        pivot_df.get('遲到', 0) + pivot_df.get('早退', 0)
    ).fillna(0)
    pivot_df['現金'] = pivot_df['實發淨薪'] - pivot_df['匯入銀行']

    # 欄位重新排序
    total_cols = ['實發淨薪', '應發總額', '應扣總額', '申報薪資', '匯入銀行', '現金']
    final_cols = ['員工姓名'] + total_cols + earning_cols + deduction_cols
    
    result_df = pivot_df[[c for c in final_cols if c in pivot_df.columns]]

    # 格式化所有數字欄位為整數
    for col in result_df.columns:
        if col != '員工姓名' and pd.api.types.is_numeric_dtype(result_df[col]):
            result_df[col] = result_df[col].astype(int)
            
    return result_df, item_types


# ---【新函式 1】篩選底薪低於新標準的員工 ---
def get_employees_below_minimum_wage(conn, new_minimum_wage: int):
    """
    查詢所有在職且目前底薪低於新基本工資的員工。
    """
    query = """
    WITH latest_salary AS (
        SELECT
            employee_id,
            base_salary,
            dependents,
            -- 使用 ROW_NUMBER() 找出每位員工最新的薪資紀錄
            ROW_NUMBER() OVER(PARTITION BY employee_id ORDER BY start_date DESC) as rn
        FROM salary_base_history
    )
    SELECT
        e.id as employee_id,
        e.name_ch as "員工姓名",
        ls.base_salary as "目前底薪",
        ls.dependents as "目前眷屬數"
    FROM employee e
    LEFT JOIN latest_salary ls ON e.id = ls.employee_id
    WHERE
        e.resign_date IS NULL -- 只找出在職員工
        AND ls.rn = 1 -- 確保只用最新的薪資紀錄來比較
        AND ls.base_salary < ? -- 核心條件：底薪低於新標準
    ORDER BY e.id;
    """
    return pd.read_sql_query(query, conn, params=(new_minimum_wage,))

# ---【新函式 2】批次更新員工底薪 ---
def batch_update_basic_salary(conn, preview_df: pd.DataFrame, new_wage: int, effective_date):
    """
    為預覽列表中的所有員工，批次新增一筆調薪紀錄。
    """
    cursor = conn.cursor()
    try:
        data_to_insert = []
        for _, row in preview_df.iterrows():
            data_to_insert.append((
                row['employee_id'],
                new_wage, # 新的底薪
                row['目前眷屬數'], # 眷屬數維持不變
                effective_date.strftime('%Y-%m-%d'),
                None, # 結束日為空
                f"配合 {effective_date.year} 年基本工資調整" # 自動產生備註
            ))
        
        sql = """
        INSERT INTO salary_base_history
        (employee_id, base_salary, dependents, start_date, end_date, note)
        VALUES (?, ?, ?, ?, ?, ?)
        """
        cursor.executemany(sql, data_to_insert)
        conn.commit()
        return cursor.rowcount
    except Exception as e:
        conn.rollback()
        raise e