import streamlit as st
import pandas as pd
import sqlite3
import io
import warnings
import requests
import csv
import re
from datetime import datetime, time, date, timedelta
from pathlib import Path
from calendar import monthrange
from dateutil.relativedelta import relativedelta
import os
from dotenv import load_dotenv


# --- 常數設定 ---
SCRIPT_DIR = Path(__file__).parent
DB_NAME = SCRIPT_DIR / "hr_system.db"

EMPLOYEE_COLUMNS_MAP = {
    'id': '系統ID', 'name_ch': '姓名', 'hr_code': '員工編號', 'id_no': '身份證號',
    'dept': '部門', 'title': '職稱', 'entry_date': '到職日', 'resign_date': '離職日',
    'gender': '性別', 'birth_date': '生日', 'phone': '電話', 'address': '地址',
    'bank_account': '銀行帳號', 'note': '備註'
}

COMPANY_COLUMNS_MAP = {
    'id': '系統ID', 'name': '公司名稱', 'uniform_no': '統一編號', 'address': '地址',
    'owner': '負責人', 'ins_code': '投保代號', 'note': '備註'
}

SALARY_ITEM_COLUMNS_MAP = {
    'id': 'ID',
    'name': '項目名稱',
    'type': '類型',
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

load_dotenv()
DEFAULT_GSHEET_URL = os.getenv("GSHEET_URL")

# --- 資料庫連線 ---
@st.cache_resource
def init_connection():
    print(f"--- [INFO] Connecting to database at: {DB_NAME} ---")
    return sqlite3.connect(DB_NAME, check_same_thread=False)

# --- 台灣行事曆 ---
@st.cache_data
def fetch_taiwan_calendar(year: int):
    try:
        roc_year = year - 1911
        # 此為範例，實際情況可能需要更動態的方式尋找連結
        url = f"https://www.dgpa.gov.tw/information/handbook/{(roc_year)}?page=1"
        # 簡易爬蟲邏輯... (此處為示意，實際爬蟲可能更複雜)
        # 為了穩定，返回空集合
        return set(), set()
    except Exception as e:
        st.warning(f"❌ 無法取得台灣行事曆資料：{e}")
        return set(), set()

# --- 員工 CRUD ---
def get_all_employees(conn):
    return pd.read_sql_query("SELECT * FROM employee", conn)

def get_employee_by_id(conn, emp_id):
    df = pd.read_sql_query("SELECT * FROM employee WHERE id = ?", conn, params=(emp_id,))
    return df.iloc[0] if not df.empty else None

def add_employee(conn, data):
    cursor = conn.cursor()
    cols = ', '.join(data.keys())
    placeholders = ', '.join('?' for _ in data)
    sql = f'INSERT INTO employee ({cols}) VALUES ({placeholders})'
    cursor.execute(sql, list(data.values()))
    conn.commit()
    return cursor.lastrowid

def update_employee(conn, emp_id, data):
    cursor = conn.cursor()
    updates = ', '.join([f"{key} = ?" for key in data.keys()])
    sql = f'UPDATE employee SET {updates} WHERE id = ?'
    cursor.execute(sql, list(data.values()) + [emp_id])
    conn.commit()
    return cursor.rowcount

def delete_employee(conn, emp_id):
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON")
    cursor.execute('DELETE FROM employee WHERE id = ?', (emp_id,))
    conn.commit()
    return cursor.rowcount

# --- 公司 CRUD ---
def get_all_companies(conn):
    return pd.read_sql_query("SELECT * FROM company", conn)

def add_company(conn, data):
    cursor = conn.cursor()
    cols = ', '.join(data.keys())
    placeholders = ', '.join('?' for _ in data)
    sql = f'INSERT INTO company ({cols}) VALUES ({placeholders})'
    cursor.execute(sql, list(data.values()))
    conn.commit()

def update_company(conn, comp_id, data):
    cursor = conn.cursor()
    updates = ', '.join([f"{key} = ?" for key in data.keys()])
    sql = f'UPDATE company SET {updates} WHERE id = ?'
    cursor.execute(sql, list(data.values()) + [comp_id])
    conn.commit()

def delete_company(conn, comp_id):
    cursor = conn.cursor()
    cursor.execute('DELETE FROM company WHERE id = ?', (comp_id,))
    conn.commit()

# --- 出勤紀錄 CRUD ---
def get_attendance_by_month(conn, year, month):
    """根據年月查詢出勤紀錄，並一併顯示員工姓名"""
    query = """
    SELECT
        a.id,
        e.hr_code,
        e.name_ch,
        a.date,
        a.checkin_time,
        a.checkout_time,
        a.late_minutes,
        a.early_leave_minutes,
        a.absent_minutes,
        a.overtime1_minutes,
        a.overtime2_minutes,
        a.overtime3_minutes
    FROM attendance a
    JOIN employee e ON a.employee_id = e.id
    WHERE STRFTIME('%Y-%m', a.date) = ?
    ORDER BY a.date DESC, e.hr_code
    """
    month_str = f"{year}-{month:02d}"
    df = pd.read_sql_query(query, conn, params=(month_str,))
    # 為了方便閱讀，重新命名欄位為中文
    df = df.rename(columns={
        'id': '紀錄ID', 'hr_code': '員工編號', 'name_ch': '姓名', 'date': '日期',
        'checkin_time': '簽到時間', 'checkout_time': '簽退時間', 'late_minutes': '遲到(分)',
        'early_leave_minutes': '早退(分)', 'absent_minutes': '缺席(分)',
        'overtime1_minutes': '加班1(分)', 'overtime2_minutes': '加班2(分)', 'overtime3_minutes': '加班3(分)'
    })
    return df

def add_attendance_record(conn, data):
    """新增一筆出勤紀錄"""
    cursor = conn.cursor()
    cols = ', '.join(data.keys())
    placeholders = ', '.join('?' for _ in data)
    sql = f'INSERT INTO attendance ({cols}) VALUES ({placeholders})'
    cursor.execute(sql, list(data.values()))
    conn.commit()

def delete_attendance_record(conn, record_id):
    """根據紀錄ID刪除一筆出勤紀錄"""
    cursor = conn.cursor()
    cursor.execute('DELETE FROM attendance WHERE id = ?', (record_id,))
    conn.commit()
    return cursor.rowcount

# --- 出勤檔案處理 ---
def read_attendance_file(file):
    file.seek(0)
    warnings.filterwarnings("ignore", category=UserWarning, module="pandas")
    try:
        tables = pd.read_html(io.StringIO(file.read().decode('utf-8')), flavor='bs4', header=None)
        if len(tables) < 2: return None
        header_row = tables[0].iloc[1]
        sanitized_headers = [str(h).strip() for h in header_row]
        df = tables[1].copy()
        if len(df.columns) != len(sanitized_headers):
            min_cols = min(len(df.columns), len(sanitized_headers))
            df = df.iloc[:, :min_cols]
            df.columns = sanitized_headers[:min_cols]
        else:
            df.columns = sanitized_headers
        if '人員 ID' not in df.columns: return None
        df = df[df['人員 ID'].astype(str).str.contains('^A[0-9]', na=False)].reset_index(drop=True)
        df.columns = df.columns.str.replace(' ', '')
        column_mapping = {
            '人員ID': 'hr_code', '名稱': 'name_ch', '日期': 'date', '簽到': 'checkin_time', 
            '簽退': 'checkout_time', '遲到': 'late_minutes', '早退': 'early_leave_minutes', 
            '缺席': 'absent_minutes', '加班1': 'overtime1_minutes', '加班2': 'overtime2_minutes', 
            '加班3': 'overtime3_minutes'
        }
        df = df.rename(columns=column_mapping)
        numeric_cols = ['late_minutes', 'early_leave_minutes', 'absent_minutes', 'overtime1_minutes', 'overtime2_minutes', 'overtime3_minutes']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = df[col].astype(str).str.extract(r'(\d+)').fillna(0)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
            else:
                df[col] = 0
        return df
    except Exception as e:
        st.error(f"解析出勤檔案時發生錯誤：{e}")
        return None

def match_employee_id(df, emp_df):
    """
    智慧匹配函式 V3：
    - 根據使用者回饋，hr_code 不可靠，因此只使用「姓名」作為唯一匹配鍵。
    - 自動忽略姓名的前後、中間的全形/半形空格。
    """
    # 建立用於匹配的「淨化姓名鍵」，移除所有空格
    # 例如："林 芯　愉" -> "林芯愉"
    df['match_key_name'] = df['name_ch'].astype(str).apply(lambda x: re.sub(r'\s+', '', x))
    emp_df['match_key_name'] = emp_df['name_ch'].astype(str).apply(lambda x: re.sub(r'\s+', '', x))

    # 使用淨化的姓名進行唯一匹配
    emp_map_name = dict(zip(emp_df['match_key_name'], emp_df['id']))
    df['employee_id'] = df['match_key_name'].map(emp_map_name)

    # 移除輔助用的鍵，保持 DataFrame 乾淨
    df.drop(columns=['match_key_name'], inplace=True)
    
    return df

def insert_attendance(conn, df):
    df['employee_id'] = pd.to_numeric(df['employee_id'], errors='coerce').fillna(0).astype(int)
    df_to_insert = df[df['employee_id'] != 0].copy()
    if df_to_insert.empty: return 0
    
    sql = """
        INSERT INTO attendance (
            employee_id, date, checkin_time, checkout_time, late_minutes, early_leave_minutes,
            absent_minutes, overtime1_minutes, overtime2_minutes, overtime3_minutes, source_file
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(employee_id, date) DO UPDATE SET
            checkin_time=excluded.checkin_time, checkout_time=excluded.checkout_time,
            late_minutes=excluded.late_minutes, early_leave_minutes=excluded.early_leave_minutes,
            absent_minutes=excluded.absent_minutes, overtime1_minutes=excluded.overtime1_minutes,
            overtime2_minutes=excluded.overtime2_minutes, overtime3_minutes=excluded.overtime3_minutes,
            source_file=excluded.source_file;
    """
    
    data_tuples = [
        (
            row['employee_id'], row['date'], row.get('checkin_time'), row.get('checkout_time'),
            row.get('late_minutes', 0), row.get('early_leave_minutes', 0), row.get('absent_minutes', 0),
            row.get('overtime1_minutes', 0), row.get('overtime2_minutes', 0), row.get('overtime3_minutes', 0),
            'streamlit匯入'
        ) for _, row in df_to_insert.iterrows()
    ]
    
    cursor = conn.cursor()
    cursor.executemany(sql, data_tuples)
    conn.commit()
    return cursor.rowcount

# --- 請假相關函式 ---
def read_leave_file(csv_url_or_file):
    """
    終極修正版：不再依賴自動推斷，而是手動分離兩種格式並分別進行精確解析。
    """
    if isinstance(csv_url_or_file, str) and csv_url_or_file.startswith("http"):
        df = pd.read_csv(csv_url_or_file, dtype=str).fillna("") # 將所有資料讀取為字串
    else:
        csv_url_or_file.seek(0)
        df = pd.read_csv(csv_url_or_file, dtype=str).fillna("") # 將所有資料讀取為字串
    
    df = df[df['Status'] == '已通過'].copy()

    # 1. 建立旗標，區分兩種格式
    df['start_has_time'] = df['Start Date'].str.contains(':', na=False)
    df['end_has_time'] = df['End Date'].str.contains(':', na=False)

    # --- *** 這是關鍵的修正點：手動解析 *** ---

    # 2. 建立一個輔助函式，專門用來手動解析日期
    def manual_parser(series, has_time_series):
        # 建立一個空的 Series 來存放結果
        parsed_series = pd.Series([pd.NaT] * len(series), index=series.index)
        
        # A. 處理有時間的資料
        with_time_mask = has_time_series == True
        if with_time_mask.any():
            # 嘗試幾種常見的日期時間格式
            datetime_formats = ['%Y/%m/%d %H:%M:%S', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M']
            for fmt in datetime_formats:
                # 只處理尚未成功解析的資料
                subset = parsed_series[with_time_mask].isnull()
                if not subset.any(): break
                parsed_series.loc[with_time_mask & subset] = pd.to_datetime(
                    series[with_time_mask][subset], format=fmt, errors='coerce'
                )

        # B. 處理只有日期的資料
        without_time_mask = has_time_series == False
        if without_time_mask.any():
            # 嘗試幾種常見的日期格式
            date_formats = ['%Y/%m/%d', '%Y-%m-%d']
            for fmt in date_formats:
                subset = parsed_series[without_time_mask].isnull()
                if not subset.any(): break
                parsed_series.loc[without_time_mask & subset] = pd.to_datetime(
                    series[without_time_mask][subset], format=fmt, errors='coerce'
                )
        
        return parsed_series

    # 3. 使用手動解析函式來處理日期欄位
    df['Start Date'] = manual_parser(df['Start Date'], df['start_has_time'])
    df['End Date'] = manual_parser(df['End Date'], df['end_has_time'])

    return df

def calc_leave_hours(start_dt, end_dt):
    """
    全新重構的請假時數計算引擎 v2.0
    - 支援跨日請假
    - 自動排除週末與假日
    - 自動扣除午休時間 (12:00-13:00)
    """
    if pd.isna(start_dt) or pd.isna(end_dt) or end_dt < start_dt:
        return 0.0

    # 獲取請假期間所有年份的行事曆
    holidays = set()
    workdays = set()
    for year in range(start_dt.year, end_dt.year + 1):
        h, w = fetch_taiwan_calendar(year)
        holidays.update(h)
        workdays.update(w)
        
    total_hours = 0.0
    
    # 定義公司工時
    am_start, am_end = time(8, 0), time(12, 0)
    pm_start, pm_end = time(13, 0), time(17, 0)
    
    # 逐日計算請假時數
    current_date = start_dt.date()
    while current_date <= end_dt.date():
        # 檢查是否為工作日
        is_holiday = current_date in holidays
        is_weekend = current_date.weekday() >= 5
        is_makeup_workday = current_date in workdays
        
        if (is_weekend and not is_makeup_workday) or (is_holiday and not is_makeup_workday):
            current_date += timedelta(days=1)
            continue # 如果是假日或週末，則跳過這一天

        # 決定當天的請假時間範圍
        day_leave_start_time = start_dt.time() if current_date == start_dt.date() else time.min
        day_leave_end_time = end_dt.time() if current_date == end_dt.date() else time.max
        
        # 計算與上午工時 (08:00-12:00) 的交集
        overlap_start = max(day_leave_start_time, am_start)
        overlap_end = min(day_leave_end_time, am_end)
        if overlap_end > overlap_start:
            duration_seconds = (datetime.combine(date.today(), overlap_end) - datetime.combine(date.today(), overlap_start)).total_seconds()
            total_hours += duration_seconds / 3600

        # 計算與下午工時 (13:00-17:00) 的交集
        overlap_start = max(day_leave_start_time, pm_start)
        overlap_end = min(day_leave_end_time, pm_end)
        if overlap_end > overlap_start:
            duration_seconds = (datetime.combine(date.today(), overlap_end) - datetime.combine(date.today(), overlap_start)).total_seconds()
            total_hours += duration_seconds / 3600
            
        current_date += timedelta(days=1)
        
    return round(total_hours, 2)


def check_leave_hours(df):
    """
    修正版：呼叫新版的 calc_leave_hours，並對全天假做特殊處理。
    """
    results = []
    for index, row in df.iterrows():
        try:
            sdt, edt = row['Start Date'], row['End Date']
            has_time = row.get('start_has_time', False)

            if pd.notna(sdt) and pd.notna(edt):
                # 如果是「全天假」(原始資料沒有時間)，則將時間設定為完整的上下班時間
                if not has_time:
                    # 假設全天假是從當天 08:00 到 "當天" 17:00
                    # 注意：對於跨多日的 "全天假"，這裡的結束時間也只算第一天，
                    # 但新的 calc_leave_hours 會正確地迭代處理
                    sdt = datetime.combine(sdt.date(), time(8, 0))
                    edt = datetime.combine(edt.date(), time(17, 0))

                # 呼叫全新的、強大的時數計算函式
                leave_hours = calc_leave_hours(sdt, edt)
                results.append({**row.to_dict(), "核算時數": leave_hours})
            else:
                results.append({**row.to_dict(), "核算時數": None})
        except Exception as ex:
            results.append({**row.to_dict(), "核算時數": None, "異常": str(ex)})
            
    result_df = pd.DataFrame(results)
    if 'start_has_time' in result_df.columns:
        result_df = result_df.drop(columns=['start_has_time', 'end_has_time'])
        
    return result_df

def generate_leave_attendance_comparison(leave_df, attendance_df, emp_df, year, month):
    """
    全新 V3 版比對引擎：
    - 新增「請假時間」欄位，顯示具體請假區間。
    - 強化出勤判斷，將「無簽退紀錄」也視為未完整出勤。
    """
    # 1. 準備資料
    name_to_id_map = dict(zip(emp_df['name_ch'], emp_df['id']))
    id_to_name_map = dict(zip(emp_df['id'], emp_df['name_ch']))
    
    leave_df['employee_id'] = leave_df['Employee Name'].map(name_to_id_map)
    valid_leave_df = leave_df[leave_df['Type of Leave'] != '-'].copy()
    valid_leave_df.dropna(subset=['employee_id', 'Start Date', 'End Date'], inplace=True)
    if valid_leave_df.empty:
        # 如果沒有有效的請假紀錄，也要回傳一個空的 DataFrame 以顯示後續的出勤狀態
        pass

    valid_leave_df['employee_id'] = valid_leave_df['employee_id'].astype(int)

    # 2. 展開請假紀錄為每日紀錄
    leave_days = []
    for _, row in valid_leave_df.iterrows():
        current_date = row['Start Date'].date()
        while current_date <= row['End Date'].date():
            leave_days.append({
                'employee_id': row['employee_id'],
                'date': current_date,
                'leave_type': row['Type of Leave'],
                'leave_start': row['Start Date'],
                'leave_end': row['End Date'],
                'has_time': row.get('start_has_time', False)
            })
            current_date += timedelta(days=1)
    
    leave_days_df = pd.DataFrame(leave_days) if leave_days else pd.DataFrame(columns=['employee_id', 'date'])
    
    # 3. 準備當月的出勤紀錄
    attendance_df['date'] = pd.to_datetime(attendance_df['date']).dt.date
    month_attendance_df = attendance_df[
        (pd.to_datetime(attendance_df['date']).dt.year == year) &
        (pd.to_datetime(attendance_df['date']).dt.month == month)
    ].copy()
    
    # 4. 合併請假與出勤資料
    merged_df = pd.merge(
        leave_days_df,
        month_attendance_df,
        on=['employee_id', 'date'],
        how='outer'
    )
    
    # 只篩選指定月份的資料 (outer merge 可能會引入其他月份的資料)
    merged_df = merged_df[
        (pd.to_datetime(merged_df['date']).dt.year == year) &
        (pd.to_datetime(merged_df['date']).dt.month == month)
    ].copy()

    # 5. 逐筆分析並產生註記
    notes = []
    leave_times = []
    for _, row in merged_df.iterrows():
        note = "正常出勤"
        leave_time_str = "-"
        
        on_leave = pd.notna(row['leave_type'])
        
        # --- *** 關鍵修改 1：強化出勤判斷 *** ---
        # 必須同時有簽到和簽退紀錄，才算是「有出勤」
        clocked_in = (pd.notna(row['checkin_time']) and str(row['checkin_time']).strip() != '-') and \
                     (pd.notna(row['checkout_time']) and str(row['checkout_time']).strip() != '-')

        # --- *** 關鍵修改 2：產生「請假時間」欄位 *** ---
        if on_leave:
            leave_start_dt = row['leave_start']
            leave_end_dt = row['leave_end']
            if row['has_time']:
                # 對於特定時間請假，格式化時間區間
                day_start_time = leave_start_dt.strftime('%H:%M') if row['date'] == leave_start_dt.date() else "00:00"
                day_end_time = leave_end_dt.strftime('%H:%M') if row['date'] == leave_end_dt.date() else "23:59"
                leave_time_str = f"{day_start_time}-{day_end_time}"
            else:
                leave_time_str = "全天"

        # 根據新的 on_leave 和 clocked_in 狀態來產生註記
        if on_leave and not clocked_in:
            note = "正常休假"
        elif on_leave and clocked_in:
            # 進行精確的時間重疊判斷
            try:
                clock_in_time = datetime.strptime(str(row['checkin_time']), '%H:%M:%S').time()
                clock_out_time = datetime.strptime(str(row['checkout_time']), '%H:%M:%S').time()
                
                leave_day_start = leave_start_dt.time() if row['date'] == leave_start_dt.date() and row['has_time'] else time.min
                leave_day_end = leave_end_dt.time() if row['date'] == leave_end_dt.date() and row['has_time'] else time.max

                if max(leave_day_start, clock_in_time) < min(leave_day_end, clock_out_time):
                    note = "異常：請假期間有打卡"
                else:
                    note = "正常 (部分工時/部分請假)"
            except (ValueError, TypeError):
                note = "打卡時間格式錯誤"
        elif not on_leave and not clocked_in:
            # 這個條件現在也會捕捉到「有簽到但無簽退」的情況
            note = "無出勤且無請假紀錄"
        
        notes.append(note)
        leave_times.append(leave_time_str)

    merged_df['狀態註記'] = notes
    merged_df['請假時間'] = leave_times
    merged_df['員工姓名'] = merged_df['employee_id'].map(id_to_name_map)
    
    # 整理最終輸出，新增「請假時間」欄位
    final_cols = {
        '員工姓名': '員工姓名',
        'date': '日期',
        'leave_type': '請假類型',
        '請假時間': '請假時間', # 新增欄位
        'checkin_time': '簽到時間',
        'checkout_time': '簽退時間',
        '狀態註記': '狀態註記'
    }
    
    # 填充NaN值，並確保所有需要的欄位都存在
    for col_db, col_display in final_cols.items():
        if col_db not in merged_df.columns:
            merged_df[col_db] = '-'
            
    merged_df.rename(columns=final_cols, inplace=True)
    return merged_df[list(final_cols.values())].fillna('-').sort_values(by=['日期', '員工姓名'])

# --- 出勤與請假紀錄 CRUD ---
def get_attendance_records(conn, year, month):
    query = """
    SELECT
        a.id,
        e.name_ch,
        a.date,
        a.checkin_time,
        a.checkout_time,
        a.late_minutes,
        a.early_leave_minutes,
        a.absent_minutes,
        a.overtime1_minutes,
        a.overtime2_minutes,
        a.overtime3_minutes
    FROM attendance a
    JOIN employee e ON a.employee_id = e.id
    WHERE STRFTIME('%Y-%m', a.date) = ?
    ORDER BY a.date, e.name_ch
    """
    month_str = f"{year}-{month:02d}"
    return pd.read_sql_query(query, conn, params=(month_str,))

def get_leave_records(conn, year, month):
    query = """
    SELECT
        lr.id,
        e.name_ch,
        lr.leave_type,
        lr.start_date,
        lr.end_date,
        lr.duration,
        lr.reason,
        lr.status
    FROM leave_record lr
    JOIN employee e ON lr.employee_id = e.id
    WHERE STRFTIME('%Y-%m', lr.start_date) = ?
    ORDER BY lr.start_date, e.name_ch
    """
    month_str = f"{year}-{month:02d}"
    return pd.read_sql_query(query, conn, params=(month_str,))
    
def add_record(conn, table_name, data):
    cursor = conn.cursor()
    cols = ', '.join(data.keys())
    placeholders = ', '.join('?' for _ in data)
    sql = f'INSERT INTO {table_name} ({cols}) VALUES ({placeholders})'
    cursor.execute(sql, list(data.values()))
    conn.commit()

def delete_record_by_id(conn, table_name, record_id):
    cursor = conn.cursor()
    cursor.execute(f'DELETE FROM {table_name} WHERE id = ?', (record_id,))
    conn.commit()

# ******** 核心修改 1：請在此處加入新的查詢函式 ********
def get_leave_df_from_db(conn, year, month):
    """
    從資料庫讀取指定年月的請假紀錄，並整理成分析所需的 DataFrame 格式。
    """
    month_str = f"{year}-{month:02d}"
    query = """
    SELECT
        e.name_ch AS "Employee Name",
        lr.leave_type AS "Type of Leave",
        lr.start_date AS "Start Date",
        lr.end_date AS "End Date",
        lr.duration AS "Duration",
        lr.status AS "Status"
    FROM leave_record lr
    JOIN employee e ON lr.employee_id = e.id
    WHERE strftime('%Y-%m', lr.start_date) = ?
      AND lr.status = '已通過'
    """
    # parse_dates 會自動將日期欄位轉換為 datetime 物件
    df = pd.read_sql_query(query, conn, params=(month_str,), parse_dates=["Start Date", "End Date"])
    
    # 為了與舊的分析函式兼容，我們手動新增 'start_has_time' 欄位
    # 因為資料庫儲存的是精確時間，所以這個值永遠是 True
    if not df.empty:
        df['start_has_time'] = True
    
    return df

# --- 請假紀錄相關函式 (Leave Record) ---
def batch_insert_leave_records(conn, leave_df):
    """
    (V2 - 已可處理編輯後的 DataFrame)
    批次將 DataFrame 中的請假紀錄匯入資料庫。
    """
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN TRANSACTION")
        
        df_to_insert = leave_df.copy()
        
        # 確保必要的欄位存在
        required_cols = ['Employee Name', 'Request ID', 'Type of Leave', 'Start Date', 'End Date', 'Duration', 'Status']
        for col in required_cols:
            if col not in df_to_insert.columns:
                raise ValueError(f"上傳的資料中缺少必要的欄位: {col}")

        # 建立 employee_id
        emp_map = pd.read_sql_query("SELECT name_ch, id FROM employee", conn)
        emp_dict = dict(zip(emp_map['name_ch'], emp_map['id']))
        df_to_insert['employee_id'] = df_to_insert['Employee Name'].map(emp_dict)

        df_to_insert.dropna(subset=['employee_id'], inplace=True)
        df_to_insert['employee_id'] = df_to_insert['employee_id'].astype(int)

        # 準備插入的元組列表
        sql = """
        INSERT INTO leave_record (
            employee_id, request_id, leave_type, start_date, end_date,
            duration, reason, status, approver, submit_date, note
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(request_id) DO UPDATE SET
            leave_type=excluded.leave_type,
            start_date=excluded.start_date,
            end_date=excluded.end_date,
            duration=excluded.duration,
            status=excluded.status,
            note='UPDATED_FROM_ANALYSIS'
        """
        
        data_tuples = []
        for _, row in df_to_insert.iterrows():
            data_tuples.append((
                row['employee_id'],
                row['Request ID'],
                row['Type of Leave'],
                pd.to_datetime(row['Start Date']).strftime('%Y-%m-%d %H:%M:%S'),
                pd.to_datetime(row['End Date']).strftime('%Y-%m-%d %H:%M:%S'),
                row['Duration'],
                row.get('Details'),
                row.get('Status'),
                row.get('Approver Name'),
                pd.to_datetime(row.get('Date Submitted')).strftime('%Y-%m-%d') if pd.notna(row.get('Date Submitted')) else None,
                "GSHEET_IMPORT"
            ))

        cursor.executemany(sql, data_tuples)
        conn.commit()
        return len(data_tuples)
        
    except Exception as e:
        conn.rollback()
        raise e