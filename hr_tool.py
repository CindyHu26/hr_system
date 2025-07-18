# hr_tool.py
import streamlit as st
from utils import init_connection
# 從各個頁面模組中引用 show_page 函式
from page_crud_employee import show_page as show_employee_page
from page_crud_company import show_page as show_company_page
from page_crud_attendance import show_page as show_attendance_crud_page
from page_leave_analysis import show_page as show_analysis_page
from page_salary_item import show_page as show_salary_item_page 

# 建立資料庫連線
conn = init_connection()

# 定義頁面對應的函式
PAGES = {
    "員工管理 (CRUD)": show_employee_page,
    "公司管理 (CRUD)": show_company_page,
    "出勤紀錄管理 (CRUD)": show_attendance_crud_page,
    "請假與異常分析": show_analysis_page,
    "薪資項目管理": show_salary_item_page,
}

st.sidebar.title("HRIS 人資系統")
# 建立選擇器
selection = st.sidebar.radio("請選擇功能頁面", list(PAGES.keys()))

# 根據選擇，執行對應頁面的函式
page_function = PAGES[selection]
page_function(conn)

# 可選擇在最後關閉連線 (取決於應用生命週期)
# conn.close()