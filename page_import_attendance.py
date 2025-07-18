# page_import_attendance.py (姓名匹配版)
import streamlit as st
import pandas as pd
import re
import traceback
from utils import (
    read_attendance_file,
    match_employee_id,
    insert_attendance,
    get_all_employees
)

def show_page(conn):
    st.header("打卡機出勤檔案匯入")
    st.info("系統將使用「姓名」作為唯一匹配依據，並自動忽略姓名中的所有空格。請確保打卡檔姓名與員工資料庫中的姓名一致。")
    
    uploaded_file = st.file_uploader("上傳打卡機檔案 (XLS)", type=['xls'])
    
    if uploaded_file:
        df = read_attendance_file(uploaded_file)
        
        if df is not None and not df.empty:
            st.write("---")
            st.subheader("1. 檔案解析預覽")
            st.dataframe(df.head(5))

            st.write("---")
            st.subheader("2. 員工姓名匹配")
            try:
                emp_df = get_all_employees(conn)
                if emp_df.empty:
                    st.error("資料庫中沒有員工資料，無法進行匹配。請先至「員工管理」頁面新增員工。")
                    return
                
                df_matched = match_employee_id(df, emp_df)
                
                matched_count = df_matched['employee_id'].notnull().sum()
                unmatched_count = len(df_matched) - matched_count
                
                st.info(f"匹配結果：成功 **{matched_count}** 筆 / 失敗 **{unmatched_count}** 筆。")

                if unmatched_count > 0:
                    st.error(f"有 {unmatched_count} 筆紀錄匹配失敗，將不會被匯入：")
                    
                    unmatched_df = df_matched[df_matched['employee_id'].isnull()]
                    st.dataframe(unmatched_df[['hr_code', 'name_ch', 'date']])

                    with st.expander("🔍 點此展開進階偵錯，查看失敗原因"):
                        st.warning("此工具會顯示資料的「原始樣貌」，幫助您找出例如空格、特殊字元等看不見的差異。")
                        for index, row in unmatched_df.iterrows():
                            report_name = row['name_ch']
                            report_code = row['hr_code']
                            st.markdown(f"--- \n#### 正在分析失敗紀錄: **{report_name} ({report_code})**")
                            
                            st.markdown("**打卡檔中的原始資料：**")
                            st.code(f"姓名: {report_name!r}")

                            st.markdown("**資料庫中的潛在匹配：**")
                            # 修正 AttributeError: 'Series' object has no attribute 'lower' 的錯誤
                            # 並簡化邏輯，只比對淨化後的姓名
                            emp_df['match_key_name_debug'] = emp_df['name_ch'].astype(str).apply(lambda x: re.sub(r'\s+', '', x))
                            report_name_clean = re.sub(r'\s+', '', report_name)
                            
                            potential_match_name = emp_df[emp_df['match_key_name_debug'] == report_name_clean]
                            
                            if not potential_match_name.empty:
                                st.write("依據「姓名」找到的相似資料：")
                                for _, db_row in potential_match_name.iterrows():
                                    st.code(f"姓名: {db_row['name_ch']!r}, 資料庫編號: {db_row['hr_code']!r}")
                            else:
                                st.info("在資料庫中找不到任何姓名相同的員工，請至「員工管理」頁面新增該員工。")

                st.write("---")
                st.subheader("3. 匯入資料庫")
                if st.button("確認匯入資料庫", disabled=(matched_count == 0)):
                    with st.spinner("正在寫入資料庫..."):
                        inserted_count = insert_attendance(conn, df_matched)
                    st.success(f"處理完成！成功匯入/更新了 {inserted_count} 筆出勤紀錄！")
                    st.info("注意：匯入的僅為「成功匹配」的紀錄。")

            except Exception as e:
                st.error(f"匹配或匯入過程中發生錯誤：{e}")
                st.error(traceback.format_exc())
        else:
            st.error("檔案解析失敗，請確認檔案格式是否為正確的 report.xls 檔案。")