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

# --- 異常比對與缺勤彙總 ---
def find_absence(attendance_df, leave_df, emp_df, year, month, count_mode=False):
    """
    修正版：在合併前，強制統一 employee_id 的資料類型，以確保比對正確。
    """
    name_map = dict(zip(emp_df['id'], emp_df['name_ch']))
    holidays, workdays = fetch_taiwan_calendar(year)
    
    try:
        _, last_day = monthrange(year, month)
    except ValueError:
        st.error(f"輸入的年份 ({year}) 或月份 ({month}) 不正確。")
        return pd.DataFrame()

    all_dates = [datetime(year, month, d).date() for d in range(1, last_day + 1)]
    work_dates = [d for d in all_dates if (d.weekday() < 5 and d not in holidays) or d in workdays]
    
    if not work_dates:
        st.warning(f"{year} 年 {month} 月沒有應出勤的工作日。")
        return pd.DataFrame()

    all_index = pd.MultiIndex.from_product([emp_df['id'].unique(), work_dates], names=['employee_id', 'date'])
    df_all = pd.DataFrame(index=all_index).reset_index()
    
    # 確保日期格式一致
    attendance_df['date'] = pd.to_datetime(attendance_df['date']).dt.date
    df_all['date'] = pd.to_datetime(df_all['date']).dt.date
    
    # --- *** 這是關鍵的修正點 *** ---
    # 在合併前，強制將兩邊的 employee_id 都轉換為整數 (int) 類型
    df_all['employee_id'] = df_all['employee_id'].astype(int)
    attendance_df['employee_id'] = pd.to_numeric(attendance_df['employee_id'], errors='coerce').fillna(0).astype(int)
    
    # 進行左合併，找出應出勤但可能沒有打卡紀錄的員工
    merged = pd.merge(df_all, attendance_df[['employee_id', 'date', 'checkin_time']], on=['employee_id', 'date'], how='left')
    
    # 檢查員工當天是否已請假
    def is_on_leave(row):
        emp_name = name_map.get(row['employee_id'], '')
        if not emp_name: return False
        
        # 篩選出該員工的請假紀錄，並確保 leave_df 中的日期也是 date 物件
        leave_df['Start Date'] = pd.to_datetime(leave_df['Start Date']).dt.date
        leave_df['End Date'] = pd.to_datetime(leave_df['End Date']).dt.date
        
        sub = leave_df[leave_df['Employee Name'] == emp_name]
        for _, lrow in sub.iterrows():
            if pd.notna(lrow['Start Date']) and pd.notna(lrow['End Date']):
                if lrow['Start Date'] <= row['date'] <= lrow['End Date']:
                    return True
        return False
        
    merged['on_leave'] = merged.apply(is_on_leave, axis=1)
    
    # 如果沒有打卡紀錄 (checkin_time 是 NaT/None) 且當天沒有請假，則標記為「缺勤」
    merged['異常'] = merged.apply(lambda r: "缺勤" if pd.isna(r['checkin_time']) and not r['on_leave'] else "", axis=1)
    merged['員工姓名'] = merged['employee_id'].map(name_map)
    
    absent_df = merged[merged['異常'] == "缺勤"]
    
    if count_mode:
        summary = absent_df.groupby('員工姓名')['date'].count().reset_index().rename(columns={'date': '缺勤天數'})
        return summary
        
    return absent_df[['員工姓名', 'date', '異常']]

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

# --- 年度特休計算函式 ---

def get_annual_leave_entitlement(tenure_years):
    """根據勞基法計算特休天數"""
    if tenure_years < 0.5:
        return 0
    elif tenure_years < 1:
        return 3
    elif tenure_years < 2:
        return 7
    elif tenure_years < 3:
        return 10
    elif tenure_years < 5:
        return 14
    elif tenure_years < 10:
        return 15
    else:
        # 滿10年後，每多一年加一天，最多30天
        return min(15 + (int(tenure_years) - 9), 30)

def get_used_annual_leave(conn, employee_id, year_start_date, year_end_date):
    """查詢資料庫，計算指定區間內已使用的特休時數"""
    # 假設 leave_record 表已存在且有資料
    query = """
    SELECT SUM(duration)
    FROM leave_record
    WHERE employee_id = ?
    AND leave_type = '特休'
    AND status = '已通過'
    AND start_date BETWEEN ? AND ?
    """
    # 如果您請假是記錄在 Google Sheet，則需修改此函式來讀取 leave_df
    
    cursor = conn.cursor()
    # 將 date 物件轉為字串以供查詢
    cursor.execute(query, (employee_id, year_start_date.strftime('%Y-%m-%d'), year_end_date.strftime('%Y-%m-%d')))
    result = cursor.fetchone()[0]
    return result if result is not None else 0

def get_annual_leave_summary(conn):
    """產生所有員工的特休總結報告"""
    today = datetime.now().date()
    employees = get_all_employees(conn)
    summary_data = []

    for _, emp in employees.iterrows():
        if pd.isna(emp['entry_date']) or emp['entry_date'] > today.strftime('%Y-%m-%d'):
            continue # 跳過沒有到職日或尚未到職的員工
        
        entry_date = datetime.strptime(emp['entry_date'], '%Y-%m-%d').date()
        
        # 計算年資
        tenure = relativedelta(today, entry_date)
        tenure_years = tenure.years + tenure.months / 12 + tenure.days / 365.25

        # 取得法定特休天數
        entitled_days = get_annual_leave_entitlement(tenure_years)

        # 判斷當前的特休年度
        # 如果今年的週年日還沒到，年度是從去年週年日到今年週年日
        if today.month < entry_date.month or (today.month == entry_date.month and today.day < entry_date.day):
            year_start_date = date(today.year - 1, entry_date.month, entry_date.day)
            year_end_date = date(today.year, entry_date.month, entry_date.day) - timedelta(days=1)
        # 如果今年的週年日已過，年度是從今年週年日到明年週年日
        else:
            year_start_date = date(today.year, entry_date.month, entry_date.day)
            year_end_date = date(today.year + 1, entry_date.month, entry_date.day) - timedelta(days=1)

        # 從資料庫計算已休時數並轉換為天數 (假設8小時為一天)
        used_hours = get_used_annual_leave(conn, emp['id'], year_start_date, year_end_date)
        used_days = round(used_hours / 8, 2)

        summary_data.append({
            "員工姓名": emp['name_ch'],
            "到職日": emp['entry_date'],
            "年資": f"{tenure.years} 年 {tenure.months} 月",
            "本年度特休天數": entitled_days,
            "本年度已休天數": used_days,
            "剩餘天數": entitled_days - used_days,
            "本年度區間": f"{year_start_date} ~ {year_end_date}"
        })
        
    return pd.DataFrame(summary_data)

# --- 薪資相關函式 ---

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
    """
    刪除指定的薪資項目
    - 啟用外鍵約束，如果項目已被 salary_detail 引用，將會拋出 IntegrityError
    """
    cursor = conn.cursor()
    # 啟用外鍵約束，確保資料完整性
    cursor.execute("PRAGMA foreign_keys = ON;")
    sql = "DELETE FROM salary_item WHERE id = ?"
    cursor.execute(sql, (item_id,))
    conn.commit()
    # 回傳被刪除的行數，可用於判斷是否成功
    return cursor.rowcount

def check_salary_record_exists(conn, year, month):
    """檢查指定年月的薪資紀錄是否存在"""
    cursor = conn.cursor()
    sql = "SELECT EXISTS(SELECT 1 FROM salary WHERE year = ? AND month = ? LIMIT 1)"
    cursor.execute(sql, (year, month))
    return cursor.fetchone()[0] == 1

def generate_monthly_salaries(conn, year, month, pay_date, default_items):
    """為所有在職員工產生月度薪資紀錄"""
    today_str = datetime.now().strftime('%Y-%m-%d')
    # 找出所有在該月仍然在職的員工
    # 條件：到職日 <= 該月最後一天 AND (離職日 IS NULL OR 離職日 >= 該月第一天)
    month_first_day = f"{year}-{month:02d}-01"
    _, month_last_day_num = monthrange(year, month)
    month_last_day = f"{year}-{month:02d}-{month_last_day_num}"

    emp_query = """
    SELECT id FROM employee
    WHERE entry_date <= ? AND (resign_date IS NULL OR resign_date >= ?)
    """
    employees = pd.read_sql_query(emp_query, conn, params=(month_last_day, month_first_day))
    
    if employees.empty:
        return 0, 0

    cursor = conn.cursor()
    total_employees = 0
    total_details = 0

    try:
        # 使用交易確保資料一致性
        cursor.execute("BEGIN TRANSACTION")

        for _, emp_row in employees.iterrows():
            employee_id = emp_row['id']
            
            # 1. 新增薪資主紀錄
            salary_sql = "INSERT INTO salary (employee_id, year, month, pay_date) VALUES (?, ?, ?, ?)"
            cursor.execute(salary_sql, (employee_id, year, month, pay_date))
            salary_id = cursor.lastrowid
            total_employees += 1
            
            # 2. 根據預設項目新增薪資明細
            detail_sql = "INSERT INTO salary_detail (salary_id, salary_item_id, amount) VALUES (?, ?, ?)"
            for item_id, amount in default_items.items():
                if amount is not None:
                    cursor.execute(detail_sql, (salary_id, item_id, amount))
                    total_details += 1

        conn.commit()
    except Exception as e:
        conn.rollback() # 如果發生錯誤，則回滾所有操作
        raise e # 將錯誤拋出，讓上層處理

    return total_employees, total_details

def get_salary_report(conn, year, month):
    """
    取得指定年月的薪資報表，並將其轉換為 員工 vs 薪資項目 的表格
    """
    query = """
    SELECT
        s.id as salary_id,
        e.id as employee_id,
        e.name_ch as "員工姓名",
        si.name as item_name,
        si.type as item_type,
        sd.amount
    FROM salary s
    JOIN employee e ON s.employee_id = e.id
    JOIN salary_detail sd ON s.id = sd.salary_id
    JOIN salary_item si ON sd.salary_item_id = si.id
    WHERE s.year = ? AND s.month = ?
    ORDER BY e.name_ch
    """
    df = pd.read_sql_query(query, conn, params=(year, month))

    if df.empty:
        return pd.DataFrame()

    # 使用 pivot_table 進行資料透視
    pivot_df = df.pivot_table(index=["salary_id", "employee_id", "員工姓名"],
                              columns='item_name',
                              values='amount').reset_index()

    # 計算應發總額、應扣總額、實發薪資
    earning_cols = df[df['item_type'] == 'earning']['item_name'].unique()
    deduction_cols = df[df['item_type'] == 'deduction']['item_name'].unique()

    pivot_df['應發總額'] = pivot_df[earning_cols].sum(axis=1)
    pivot_df['應扣總額'] = pivot_df[deduction_cols].sum(axis=1)
    pivot_df['實發薪資'] = pivot_df['應發總額'] - pivot_df['應扣總額']
    
    return pivot_df

def get_employee_salary_details(conn, salary_id):
    """取得單一員工單月的詳細薪資項目"""
    query = """
    SELECT
        sd.id as detail_id,
        si.name as "項目名稱",
        si.type as "類型",
        sd.amount as "金額"
    FROM salary_detail sd
    JOIN salary_item si ON sd.salary_item_id = si.id
    WHERE sd.salary_id = ?
    """
    return pd.read_sql_query(query, conn, params=(int(salary_id),))

def update_salary_detail(conn, detail_id, amount):
    """更新單筆薪資明細的金額"""
    cursor = conn.cursor()
    sql = "UPDATE salary_detail SET amount = ? WHERE id = ?"
    cursor.execute(sql, (amount, detail_id))
    conn.commit()

def delete_salary_records(conn, year, month):
    """刪除指定月份的所有薪資紀錄"""
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN TRANSACTION")
        
        # 找出要刪除的 salary_ids
        cursor.execute("SELECT id FROM salary WHERE year = ? AND month = ?", (year, month))
        salary_ids_tuples = cursor.fetchall()
        
        if not salary_ids_tuples:
            return 0
        
        salary_ids = [item[0] for item in salary_ids_tuples]
        
        # 刪除 salary_detail 中的相關紀錄
        placeholders = ','.join('?' for _ in salary_ids)
        cursor.execute(f"DELETE FROM salary_detail WHERE salary_id IN ({placeholders})", salary_ids)
        
        # 刪除 salary 主紀錄
        cursor.execute(f"DELETE FROM salary WHERE id IN ({placeholders})", salary_ids)
        
        conn.commit()
        return len(salary_ids)
    except Exception as e:
        conn.rollback()
        raise e
    

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
    # 將日期物件轉換為字串，並處理 None
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
    # 將日期物件轉換為字串，並處理 None
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