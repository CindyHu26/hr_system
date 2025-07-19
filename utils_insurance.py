# utils_insurance.py
import pandas as pd
import io

def get_insurance_grades(conn):
    """取得所有勞健保級距資料"""
    return pd.read_sql_query("SELECT * FROM insurance_grade ORDER BY start_date DESC, type, grade", conn)

def parse_labor_insurance_excel(file_obj):
    """(V11 - 加入去重機制) 根據使用者提供的精準行號和欄位邏輯，解析官方 Excel 檔案"""
    try:
        df = pd.read_excel(file_obj, header=None, engine='xlrd')
        grade_row_data, salary_row_data, fee_row_data = df.iloc[36], df.iloc[37], df.iloc[68]
        start_col_index = next((i for i, text in enumerate(grade_row_data) if isinstance(text, str) and "第1級" in text), -1)
        
        if start_col_index == -1:
            raise ValueError("在第37列中找不到 '第1級'，無法定位全時勞工級距表。")

        records = []
        for i in range(start_col_index, len(salary_row_data)):
            salary, grade_text = salary_row_data.get(i), grade_row_data.get(i)
            if pd.notna(salary) and isinstance(salary, (int, float)) and isinstance(grade_text, str):
                try:
                    records.append({
                        'grade': int(''.join(filter(str.isdigit, grade_text))),
                        'salary_max': salary,
                        'employee_fee': fee_row_data.get(i),
                        'employer_fee': fee_row_data.get(i + 1),
                    })
                except (ValueError, TypeError):
                    continue
        
        if not records:
            raise ValueError("無法從指定的行號中提取有效的級距資料。請確認檔案格式未變。")

        df_final = pd.DataFrame(records).dropna(subset=['grade', 'salary_max']).drop_duplicates(subset=['grade'], keep='first')
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

        rename_map = {df.columns[0]: 'grade', df.columns[1]: 'salary_max', df.columns[2]: 'employee_fee', df.columns[6]: 'employer_fee', df.columns[7]: 'gov_fee'}
        df.rename(columns=rename_map, inplace=True)
        
        for col in ['grade', 'salary_max', 'employee_fee', 'employer_fee', 'gov_fee']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[\s,元$]', '', regex=True), errors='coerce')

        df.dropna(subset=['grade', 'salary_max'], inplace=True)
        df.drop_duplicates(subset=['grade'], keep='first', inplace=True)
        df['salary_min'] = df['salary_max'].shift(1).fillna(0) + 1
        df.loc[df.index[0], 'salary_min'] = 0
        return df[[col for col in ['grade', 'salary_min', 'salary_max', 'employee_fee', 'employer_fee', 'gov_fee'] if col in df.columns]]
    except Exception as e:
        raise ValueError(f"解析 HTML 時發生錯誤: {e}")

def batch_insert_insurance_grades(conn, df, grade_type, start_date):
    """批次插入勞健保級距資料"""
    cursor = conn.cursor()
    try:
        start_date_str = start_date.strftime('%Y-%m-%d')
        cursor.execute("BEGIN TRANSACTION")
        cursor.execute("DELETE FROM insurance_grade WHERE type = ? AND start_date = ?", (grade_type, start_date_str))
        required_cols = {'grade': 0, 'salary_min': 0, 'salary_max': 0, 'employee_fee': None, 'employer_fee': None, 'gov_fee': None, 'note': None}
        for col, default in required_cols.items():
            if col not in df.columns: df[col] = default
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
        conn.rollback(); raise e

def update_insurance_grade(conn, record_id, data):
    """更新單筆級距資料"""
    cursor = conn.cursor()
    sql = "UPDATE insurance_grade SET salary_min = ?, salary_max = ?, employee_fee = ?, employer_fee = ?, gov_fee = ?, note = ? WHERE id = ?"
    cursor.execute(sql, (data['salary_min'], data['salary_max'], data['employee_fee'], data['employer_fee'], data['gov_fee'], data['note'], record_id))
    conn.commit()

def delete_insurance_grade(conn, record_id):
    """刪除單筆級距資料"""
    cursor = conn.cursor()
    sql = "DELETE FROM insurance_grade WHERE id = ?"
    cursor.execute(sql, (record_id,))
    conn.commit()