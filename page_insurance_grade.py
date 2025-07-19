# page_insurance_grade.py (已整合起算日功能)
import streamlit as st
import pandas as pd
import requests
import io
from datetime import datetime, date
from utils_salary import (
    get_insurance_grades,
    batch_insert_insurance_grades,
    update_insurance_grade,
    delete_insurance_grade,
    parse_labor_insurance_excel,
    parse_insurance_html_table
)

def show_page(conn):
    st.header("勞健保級距表管理")
    st.info("您可以在此維護不同版本的勞、健保投保級距與費用。")

    # --- 1. 顯示目前的級距表 ---
    try:
        grades_df = get_insurance_grades(conn)
        
        st.subheader("歷史級距總覽")
        if not grades_df.empty:
            # 讓使用者可以依版本 (起算日) 篩選
            versions = sorted(pd.to_datetime(grades_df['start_date']).unique(), reverse=True)
            selected_version_date = st.selectbox(
                "選擇要檢視的版本 (依起算日)", 
                versions,
                format_func=lambda dt: dt.strftime('%Y-%m-%d')
            )
            
            display_df = grades_df[pd.to_datetime(grades_df['start_date']) == selected_version_date]
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### 勞工保險級距")
                st.dataframe(display_df[display_df['type'] == 'labor'].drop(columns=['type', 'start_date', 'id']), use_container_width=True)
            with col2:
                st.markdown("#### 全民健康保險級距")
                st.dataframe(display_df[display_df['type'] == 'health'].drop(columns=['type', 'start_date', 'id']), use_container_width=True)
        else:
            st.warning("資料庫中尚無任何級距資料，請先從下方進行更新。")

    except Exception as e:
        st.error(f"讀取級距表時發生錯誤: {e}")
        return

    st.write("---")
    st.subheader("資料更新")
    
    # --- 2. 資料更新區塊 ---
    # **核心修正：加入起算日選擇器**
    start_date = st.date_input("請選擇此份資料的「適用起算日」", value=date(datetime.now().year, 1, 1))
    
    tab1, tab2 = st.tabs(["👷 勞工保險", "🏥 全民健康保險"])

    # --- 勞保更新頁籤 ---
    with tab1:
        st.markdown("##### 更新勞工保險投保薪資分級表")
        labor_url = st.text_input("勞保局保費分攤表網址", value="https://www.bli.gov.tw/0011588.html")
        st.markdown(f"請從 [勞保局網站]({labor_url}) 下載適用於 **{start_date}** 之後的 Excel 檔案 (.xls)，並直接上傳。")
        
        uploaded_labor_file = st.file_uploader("上傳勞保級距 Excel 檔", type=['xls', 'xlsx'], key="labor_uploader")
        
        if uploaded_labor_file:
            try:
                with st.spinner("正在智慧解析您上傳的 Excel 檔案..."):
                    parsed_df = parse_labor_insurance_excel(uploaded_labor_file)
                st.success(f"成功解析檔案！此資料將以 **{start_date}** 作為起算日匯入。預覽如下：")
                st.dataframe(parsed_df)
                
                if st.button(f"✅ 確認匯入「勞保」級距表", type="primary"):
                    # **核心修正：傳入起算日**
                    count = batch_insert_insurance_grades(conn, parsed_df, 'labor', start_date)
                    st.success(f"成功匯入 {count} 筆起算日為 {start_date} 的勞保級距資料！")
                    st.rerun()
            except Exception as e:
                st.error(f"處理勞保檔案時發生錯誤：{e}")

    # --- 健保更新頁籤 ---
    with tab2:
        st.markdown("#### 更新健保投保金額分級表")
        st.markdown("系統可直接解析健保署網頁，或讓您手動上傳備用。")
        update_method = st.radio("選擇更新方式", ("從網頁自動解析 (建議)", "手動上傳檔案 (備用)"), key="health_method", horizontal=True)
        if update_method == "從網頁自動解析 (建議)":
            health_url = st.text_input("健保署保費負擔金額表網址", value="https://www.nhi.gov.tw/ch/cp-17545-f87bd-2576-1.html")
            if st.button("🔗 解析網址並預覽"):
                if not health_url:
                    st.warning("請貼上健保署的網址。")
                else:
                    try:
                        with st.spinner(f"正在從 {health_url} 下載網頁內容..."):
                            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
                            response = requests.get(health_url, headers=headers, timeout=15)
                            response.raise_for_status()
                        with st.spinner("正在解析表格內容..."):
                            st.session_state.parsed_health_df = parse_insurance_html_table(response.text)
                        st.success("成功解析健保網頁表格！")
                    except Exception as e:
                        st.error(f"處理失敗: {e}")
                        st.session_state.parsed_health_df = None
            if 'parsed_health_df' in st.session_state and st.session_state.parsed_health_df is not None:
                st.markdown("##### 解析結果預覽")
                st.dataframe(st.session_state.parsed_health_df)
                if st.button("✅ 確認將預覽資料寫入「健保」級距表", type="primary"):
                    count = batch_insert_insurance_grades(conn, st.session_state.parsed_health_df, 'health')
                    st.success(f"成功匯入 {count} 筆健保級距資料！頁面將自動刷新。")
                    del st.session_state.parsed_health_df
                    st.rerun()

    # --- 手動單筆維護 (維持原樣) ---
    with st.expander("✏️ 手動單筆微調 (適用勞健保)"):
        if not grades_df.empty:
            grades_df['display'] = (
                grades_df['type'].map({'labor': '勞保', 'health': '健保'}) + " - 第 " + 
                grades_df['grade'].astype(str) + " 級 (薪資: " + 
                grades_df['salary_min'].astype(str) + " - " + 
                grades_df['salary_max'].astype(str) + ")"
            )
            options = dict(zip(grades_df['display'], grades_df['id']))
            selected_key = st.selectbox("選擇要編輯或刪除的級距", options.keys(), index=None, placeholder="請選擇一筆紀錄...")

            if selected_key:
                record_id = options[selected_key]
                record_data = grades_df[grades_df['id'] == record_id].iloc[0]

                with st.form(f"edit_grade_{record_id}"):
                    st.markdown(f"#### 正在編輯: {selected_key}")
                    c1, c2 = st.columns(2)
                    salary_min = c1.number_input("投保薪資下限", value=int(record_data['salary_min']))
                    salary_max = c2.number_input("投保薪資上限", value=int(record_data['salary_max']))
                    
                    c3, c4, c5 = st.columns(3)
                    employee_fee = c3.number_input("員工負擔", value=int(record_data.get('employee_fee', 0) or 0))
                    employer_fee = c4.number_input("雇主負擔", value=int(record_data.get('employer_fee', 0) or 0))
                    gov_fee = c5.number_input("政府補助", value=int(record_data.get('gov_fee', 0) or 0))
                    note = st.text_input("備註", value=str(record_data.get('note', '') or ''))
                    
                    if st.form_submit_button("儲存變更", use_container_width=True):
                        new_data = {
                            'salary_min': salary_min, 'salary_max': salary_max,
                            'employee_fee': employee_fee, 'employer_fee': employer_fee,
                            'gov_fee': gov_fee, 'note': note
                        }
                        update_insurance_grade(conn, record_id, new_data)
                        st.success(f"紀錄 ID: {record_id} 已更新！")
                        st.rerun()

                if st.button("🔴 刪除此級距", key=f"delete_grade_{record_id}", type="primary"):
                    delete_insurance_grade(conn, record_id)
                    st.success(f"紀錄 ID: {record_id} 已被刪除！")
                    st.rerun()
        else:
            st.info("目前系統中沒有級距資料，請先透過上方頁籤進行更新。")