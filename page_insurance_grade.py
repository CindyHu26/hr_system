# page_insurance_grade.py
import streamlit as st
import pandas as pd
from utils import (
    get_insurance_grades,
    batch_insert_insurance_grades,
    update_insurance_grade,
    delete_insurance_grade
)

def show_page(conn):
    """
    顯示勞健保級距表管理頁面
    """
    st.header("勞健保級距表管理")
    st.info("您可以在此維護勞保與健保的投保級距與費用。建議使用批次匯入功能來更新年度資料。")

    # --- 1. 顯示目前的級距表 (Read) ---
    st.subheader("目前系統中的級距表")
    try:
        grades_df = get_insurance_grades(conn)
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 勞工保險級距")
            labor_df = grades_df[grades_df['type'] == 'labor'].drop(columns=['type'])
            st.dataframe(labor_df, use_container_width=True)
        
        with col2:
            st.markdown("#### 全民健康保險級距")
            health_df = grades_df[grades_df['type'] == 'health'].drop(columns=['type'])
            st.dataframe(health_df, use_container_width=True)

    except Exception as e:
        st.error(f"讀取級距表時發生錯誤: {e}")
        return

    st.write("---")

    # --- 2. 批次匯入 (Create/Update) ---
    with st.expander("🚀 批次匯入更新 (建議使用)", expanded=True):
        st.markdown("請上傳從[勞保局](https://www.bli.gov.tw/0014162.html)或[健保署](https://www.nhi.gov.tw/Content_List.aspx?n=556941E62735919B&topn=5FE8C9FEAE863B46)下載的級距表檔案 (CSV 或 Excel)。")
        
        upload_type = st.radio("選擇要匯入的類別", ('labor', 'health'), horizontal=True, key="upload_type")
        uploaded_file = st.file_uploader("上傳檔案", type=['csv', 'xlsx'])

        if uploaded_file:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)

                st.write("檔案預覽：")
                st.dataframe(df.head())
                
                st.warning("匯入前請確保欄位名稱與系統要求一致。")
                st.code("必要欄位: grade, salary_min, salary_max\n選填欄位: employee_fee, employer_fee, gov_fee, note")

                if st.button(f"確認匯入「{upload_type}」級距表"):
                    with st.spinner("正在清空舊資料並匯入新資料..."):
                        count = batch_insert_insurance_grades(conn, df, upload_type)
                    st.success(f"成功匯入 {count} 筆「{upload_type}」級距資料！頁面將重新整理。")
                    st.rerun()

            except Exception as e:
                st.error(f"處理上傳檔案時發生錯誤: {e}")

    # --- 3. 手動單筆維護 (Update/Delete) ---
    with st.expander("✏️ 手動單筆維護"):
        if not grades_df.empty:
            # 建立選擇器
            grades_df['display'] = (
                grades_df['type'].map({'labor': '勞保', 'health': '健保'}) + " - 第 " + 
                grades_df['grade'].astype(str) + " 級 (投保薪資: " + 
                grades_df['salary_min'].astype(str) + " - " + 
                grades_df['salary_max'].astype(str) + ")"
            )
            options = dict(zip(grades_df['display'], grades_df['id']))
            selected_key = st.selectbox("選擇要編輯或刪除的級距", options.keys(), index=None, placeholder="請選擇一筆紀錄...")

            if selected_key:
                record_id = options[selected_key]
                record_data = grades_df[grades_df['id'] == record_id].iloc[0]

                # 修改表單
                with st.form(f"edit_grade_{record_id}"):
                    st.markdown(f"#### 正在編輯: {selected_key}")
                    c1, c2 = st.columns(2)
                    salary_min = c1.number_input("投保薪資下限", value=int(record_data['salary_min']))
                    salary_max = c2.number_input("投保薪資上限", value=int(record_data['salary_max']))
                    
                    c3, c4, c5 = st.columns(3)
                    employee_fee = c3.number_input("員工負擔", value=int(record_data['employee_fee'] or 0))
                    employer_fee = c4.number_input("雇主負擔", value=int(record_data['employer_fee'] or 0))
                    gov_fee = c5.number_input("政府補助", value=int(record_data['gov_fee'] or 0))

                    note = st.text_input("備註", value=str(record_data['note'] or ''))

                    # 操作按鈕
                    update_btn, delete_btn = st.columns([1, 0.2])
                    
                    if update_btn.form_submit_button("儲存變更", use_container_width=True):
                        new_data = {
                            'salary_min': salary_min, 'salary_max': salary_max,
                            'employee_fee': employee_fee, 'employer_fee': employer_fee,
                            'gov_fee': gov_fee, 'note': note
                        }
                        update_insurance_grade(conn, record_id, new_data)
                        st.success(f"紀錄 ID: {record_id} 已更新！")
                        st.rerun()

                # 刪除按鈕放在表單外
                if st.button("🔴 刪除此級距", key=f"delete_grade_{record_id}", type="primary"):
                    delete_insurance_grade(conn, record_id)
                    st.success(f"紀錄 ID: {record_id} 已被刪除！")
                    st.rerun()
        else:
            st.info("目前系統中沒有級距資料，請先使用批次匯入功能。")