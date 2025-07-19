import streamlit as st
import pandas as pd
from datetime import datetime
from utils_salary import (
    calculate_salary_df,
    save_salary_df,
    get_previous_non_insured_names,
    batch_update_salary_details_from_excel,
    get_salary_report_for_editing,
    check_salary_records_exist,
    get_item_types,
    save_data_editor_changes
)

def show_page(conn):
    st.header("薪資單產生與管理")

    # --- 0. 初始化與年月選擇 ---
    if 'salary_workflow_step' not in st.session_state:
        st.session_state.salary_workflow_step = 'initial'
    if 'salary_draft_df' not in st.session_state:
        st.session_state.salary_draft_df = None

    c1, c2 = st.columns(2)
    today = datetime.now()
    year = c1.number_input("選擇年份", min_value=2020, max_value=today.year + 5, value=today.year)
    month = c2.number_input("選擇月份", min_value=1, max_value=12, value=today.month)

    if 'current_month' not in st.session_state or st.session_state.current_month != (year, month):
        st.session_state.salary_workflow_step = 'initial'
        st.session_state.salary_draft_df = None
        st.session_state.current_month = (year, month)

    # --- 1. 產生草稿 ---
    if st.session_state.salary_workflow_step in ['initial', 'saved']:
        st.write("---")
        records_exist = check_salary_records_exist(conn, year, month)
        
        if records_exist:
            st.success(f"✅ {year} 年 {month} 月的薪資單已儲存在資料庫中。")
        else:
            st.info(f"💡 {year} 年 {month} 月的薪資單尚未產生。")

        if st.button("🚀 產生薪資草稿 (試算)", type="primary"):
            with st.spinner("正在為所有在職員工進行薪資試算..."):
                draft_df, _ = calculate_salary_df(conn, year, month)
                st.session_state.salary_draft_df = draft_df
                st.session_state.salary_workflow_step = 'draft'
                st.rerun()

    elif st.session_state.salary_workflow_step == 'draft':
        st.write("---")
        st.subheader("📝 薪資草稿預覽與確認")
        st.warning("以下為試算結果，尚未存入資料庫。")

        draft_df = st.session_state.salary_draft_df
        if draft_df is not None and not draft_df.empty:
            
            st.markdown("##### 1. 設定勞健保自理人員")
            all_emp_names_in_draft = draft_df['員工姓名'].tolist()
            previous_non_insured = get_previous_non_insured_names(conn, year, month)
            default_selection = [name for name in previous_non_insured if name in all_emp_names_in_draft]
            
            non_insured_names = st.multiselect(
                "選擇非公司加保 (勞健保自理) 的員工",
                options=all_emp_names_in_draft,
                default=default_selection
            )

            recalculated_draft, item_types = calculate_salary_df(conn, year, month, non_insured_names=non_insured_names)
            st.session_state.salary_draft_df = recalculated_draft
            
            st.markdown("##### 2. 預覽計算結果")
            display_colored_dataframe(recalculated_draft, item_types)

            c1_btn, c2_btn, _ = st.columns([1, 1, 3])
            if c1_btn.button("✅ 確認並儲存薪資單", type="primary"):
                with st.spinner("正在將薪資單寫入資料庫..."):
                    save_salary_df(conn, year, month, recalculated_draft)
                    st.session_state.salary_workflow_step = 'saved'
                    st.session_state.salary_draft_df = None
                    st.success("薪資單已成功儲存！")
                    st.rerun()

            if c2_btn.button("❌ 放棄此草稿"):
                st.session_state.salary_workflow_step = 'initial'
                st.session_state.salary_draft_df = None
                st.rerun()

    # --- 3. 【核心修改】顯示已儲存薪資單的編輯區塊與最終預覽區塊 ---
    if st.session_state.salary_workflow_step == 'saved' or (st.session_state.salary_workflow_step == 'initial' and check_salary_records_exist(conn, year, month)):
        display_and_edit_section(conn, year, month)
        display_final_report_section(conn, year, month)


def display_and_edit_section(conn, year, month):
    """【新】顯示可互動的編輯區塊 (st.data_editor 和 Excel 上傳)"""
    st.write("---")
    st.subheader("💵 薪資明細微調 (編輯區)")
    
    with st.expander("🚀 批次上傳津貼/費用 (Excel)"):
        # ... (此區塊邏輯與前一版相同，保持不變) ...
        pass # 此處省略重複程式碼

    st.markdown("##### 手動編輯薪資項目")
    st.caption("您可以在下表中直接修改數值。修改完成後，請點擊下方的「儲存手動變更」按鈕。")
    
    report_df, item_types = get_salary_report_for_editing(conn, year, month)
    
    if not report_df.empty:
        # 準備可編輯的 DataFrame (排除總計欄位)
        all_cols = list(report_df.columns)
        total_cols = ['實發淨薪', '應發總額', '應扣總額', '申報薪資', '匯入銀行', '現金']
        editable_cols = [col for col in all_cols if col not in total_cols]
        
        # 將編輯前的資料存入 session state，以便後續比對
        st.session_state.before_edit_df = report_df[editable_cols]

        edited_df = st.data_editor(
            report_df[editable_cols],
            use_container_width=True,
            disabled=["員工姓名"],
            num_rows="dynamic",
            key=f"data_editor_{year}_{month}"
        )
        
        # 偵測是否有變更
        if not edited_df.equals(st.session_state.before_edit_df):
            st.warning("偵測到變更，請記得點擊儲存！")
            if st.button("💾 儲存手動變更", type="primary"):
                with st.spinner("正在儲存您的修改..."):
                    try:
                        save_data_editor_changes(conn, year, month, edited_df)
                        st.success("手動修改儲存成功！下方的最終預覽已更新。")
                        st.rerun()
                    except Exception as e:
                        st.error(f"儲存時發生錯誤: {e}")

def display_final_report_section(conn, year, month):
    st.write("---")
    st.subheader("📊 最終薪資單預覽 (顏色標示)")
    
    report_df, item_types = get_salary_report_for_editing(conn, year, month)
    if not report_df.empty:
        # 【核心修正】呼叫全新的 HTML 視覺化函式
        display_html_table(report_df, item_types)
    else:
        st.warning("目前資料庫中沒有可顯示的薪資紀錄。")

def display_html_table(df, item_types):
    """【V3 視覺化函式】產生一個完整的、帶有樣式的 HTML 表格，確保對齊"""
    
    # --- 準備 CSS 樣式 ---
    # 這是讓表格看起來更專業的關鍵
    table_style = """
    <style>
        .salary-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }
        .salary-table th, .salary-table td {
            border: 1px solid #e0e0e0;
            padding: 8px;
            text-align: right;
        }
        .salary-table th {
            background-color: #f7f7f9;
            font-weight: bold;
        }
        .salary-table td:first-child {
            text-align: left;
        }
    </style>
    """

    # --- 準備顏色和欄位分類 ---
    total_cols = ['實發淨薪', '應發總額', '應扣總額', '申報薪資', '匯入銀行', '現金']
    
    # --- 產生表頭 (Header) ---
    header_html = "<thead><tr>"
    for col in df.columns:
        color = "black"
        if col in total_cols: color = "blue"
        elif item_types.get(col) == 'earning': color = "green"
        elif item_types.get(col) == 'deduction': color = "red"
        header_html += f"<th><span style='color: {color};'>{col}</span></th>"
    header_html += "</tr></thead>"
    
    # --- 產生表格內容 (Body) ---
    body_html = "<tbody>"
    for index, row in df.iterrows():
        body_html += "<tr>"
        for col_name, cell_value in row.items():
            body_html += f"<td>{cell_value}</td>"
        body_html += "</tr>"
    body_html += "</tbody>"
    
    # --- 組合最終的 HTML ---
    final_html = f"{table_style}<table class='salary-table'>{header_html}{body_html}</table>"
    
    # --- 使用 st.markdown 顯示 ---
    st.markdown(final_html, unsafe_allow_html=True)