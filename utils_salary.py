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