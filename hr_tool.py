# hr_tool.py
import streamlit as st
from utils import init_connection
# 從各個頁面模組中引用 show_page 函式
from page_crud_employee import show_page as show_employee_page
from page_crud_company import show_page as show_company_page
from page_crud_attendance import show_page as show_attendance_crud_page
from page_insurance_history import show_page as show_insurance_history_page
from page_special_attendance import show_page as show_special_attendance_page
from page_leave_analysis import show_page as show_analysis_page
from page_salary_item import show_page as show_salary_item_page
from page_salary_base_history import show_page as show_salary_base_history_page
from page_insurance_grade import show_page as show_insurance_grade_page
from page_allowance_setting import show_page as show_allowance_setting_page
from page_salary_calculation import show_page as show_salary_calculation_page
from page_annual_summary import show_page as show_annual_summary_page
from page_nhi_summary import show_page as show_nhi_summary_page


# --- [核心修改] 設定頁面為寬版佈局 ---
# 這必須是第一個執行的 Streamlit 指令
st.set_page_config(layout="wide")

# 建立資料庫連線
conn = init_connection()

# --- 將頁面分組 ---
# 1. 基礎資料管理
PAGES_ADMIN = {
    "👤 員工管理": show_employee_page,
    "🏢 公司管理": show_company_page,
    "📄 員工加保管理": show_insurance_history_page,
}

# 2. 出勤與假務
PAGES_ATTENDANCE = {
    "📅 出勤紀錄管理": show_attendance_crud_page,
    "📝 特別出勤管理 (津貼加班)": show_special_attendance_page,
    "🌴 請假與異常分析": show_analysis_page,
    
}

# 3. 薪資核心功能
PAGES_SALARY = {
    "⚙️ 薪資項目管理": show_salary_item_page,
    "🏦 勞健保級距管理": show_insurance_grade_page,
    "📈 員工底薪／眷屬異動": show_salary_base_history_page,
    "➕ 員工常態薪資項設定": show_allowance_setting_page,
    "💵 薪資單產生與管理": show_salary_calculation_page,
    "📊 年度薪資總表": show_annual_summary_page,
    "健保補充保費試算": show_nhi_summary_page,
}

# 將所有頁面字典合併成一個總字典，方便後續查找
ALL_PAGES = {**PAGES_ADMIN, **PAGES_ATTENDANCE, **PAGES_SALARY}

# --- Streamlit 側邊欄 UI ---
st.sidebar.title("HRIS 人資系統")

# 建立分組的下拉選單
page_groups = {
    "基本資料管理": list(PAGES_ADMIN.keys()),
    "出勤與假務": list(PAGES_ATTENDANCE.keys()),
    "薪資核心功能": list(PAGES_SALARY.keys())
}

selected_group = st.sidebar.selectbox("選擇功能區塊", list(page_groups.keys()))

# 根據選擇的分組，顯示對應的單選按鈕
selected_page = st.sidebar.radio(
    f"--- {selected_group} ---", 
    page_groups[selected_group],
    label_visibility="collapsed"
)

# 根據最終選擇的頁面，執行對應的函式
page_function = ALL_PAGES[selected_page]
page_function(conn)