# page_salary.py
import streamlit as st
import pandas as pd
from datetime import datetime
from calendar import monthrange
from utils import (
    get_all_salary_items,
    add_salary_item,
    update_salary_item,
    check_salary_record_exists,
    generate_monthly_salaries,
    get_salary_report,
    get_employee_salary_details,
    update_salary_detail,
    delete_salary_records
)

def show_page(conn):
    st.header("💰 薪資管理")

    tab1, tab2 = st.tabs(["薪資計算與報表", "薪資項目設定"])

    # --- Tab 1: 薪資計算與報表 ---
    with tab1:
        st.subheader("月度薪資管理")
        
        c1, c2 = st.columns(2)
        today = datetime.now()
        year = c1.number_input("選擇年份", min_value=2020, max_value=today.year + 5, value=today.year, key="salary_year")
        month = c2.number_input("選擇月份", min_value=1, max_value=12, value=today.month, key="salary_month")

        # 檢查薪資紀錄是否已存在
        records_exist = check_salary_record_exists(conn, year, month)

        if records_exist:
            st.success(f"✅ {year} 年 {month} 月的薪資紀錄已產生。您可以查詢、編輯或刪除。")
            
            # --- 查詢與編輯區塊 ---
            if st.button("查詢薪資報表", key="query_salary"):
                with st.spinner("正在產生報表..."):
                    report_df = get_salary_report(conn, year, month)
                    st.session_state['salary_report'] = report_df

            if 'salary_report' in st.session_state and not st.session_state['salary_report'].empty:
                st.write("---")
                st.subheader("薪資報表")
                report_to_display = st.session_state['salary_report'].copy()
                st.dataframe(report_to_display.drop(columns=['salary_id', 'employee_id']))

                # 匯出 CSV
                fname = f"salary_report_{year}-{month:02d}.csv"
                st.download_button(
                    "下載完整報表 (CSV)",
                    report_to_display.to_csv(index=False).encode("utf-8-sig"),
                    file_name=fname
                )

                # --- 編輯單一員工薪資 ---
                with st.expander("✏️ 編輯單一員工薪資"):
                    emp_options = dict(zip(report_to_display['員工姓名'], report_to_display['salary_id']))
                    selected_emp_name = st.selectbox("選擇要編輯的員工", options=emp_options.keys())
                    
                    if selected_emp_name:
                        selected_salary_id = emp_options[selected_emp_name]
                        details_df = get_employee_salary_details(conn, selected_salary_id)
                        
                        st.write(f"正在編輯 **{selected_emp_name}** 的薪資明細：")
                        
                        for _, row in details_df.iterrows():
                            c1_edit, c2_edit = st.columns([2,1])
                            c1_edit.text(f"{row['項目名稱']} ({row['類型']})", disabled=True)
                            new_amount = c2_edit.number_input("金額", value=row['金額'], key=f"detail_{row['detail_id']}", label_visibility="collapsed")

                            if new_amount != row['金額']:
                                update_salary_detail(conn, row['detail_id'], new_amount)
                                st.success(f"已更新 {row['項目名稱']} 金額為 {new_amount}")
                                # 為了即時反應，可以提示用戶重新查詢
                                st.toast("金額已更新！請重新查詢報表以查看變更。")

            # --- 刪除紀錄 ---
            st.write("---")
            st.error("危險區域")
            if st.button(f"🔴 刪除 {year} 年 {month} 月的所有薪資紀錄", key="delete_salary"):
                if 'confirm_delete' not in st.session_state:
                    st.session_state.confirm_delete = False
                st.session_state.confirm_delete = True

            if 'confirm_delete' in st.session_state and st.session_state.confirm_delete:
                st.warning(f"您確定要永久刪除 **{year} 年 {month} 月** 的全部薪資資料嗎？此操作無法復原！")
                if st.button("我非常確定，請刪除", type="primary"):
                    try:
                        with st.spinner("正在刪除中..."):
                            count = delete_salary_records(conn, year, month)
                        st.success(f"已成功刪除 {count} 位員工的薪資紀錄。頁面將重新整理。")
                        del st.session_state.confirm_delete
                        if 'salary_report' in st.session_state:
                            del st.session_state['salary_report']
                        st.rerun()
                    except Exception as e:
                        st.error(f"刪除失敗：{e}")


        else:
            st.info(f"ℹ️ {year} 年 {month} 月的薪資紀錄尚未產生。請填寫預設值後產生。")
            
            # --- 產生紀錄區塊 ---
            with st.form("generate_salary_form"):
                st.subheader("步驟一：設定薪資預設值")
                st.warning("請為下方的薪資項目填入「預設」發放金額。您可以在產生後再對個別員工進行微調。")
                
                salary_items = get_all_salary_items(conn, active_only=True)
                default_items = {}

                st.write("**加項 (Earnings)**")
                earning_items = salary_items[salary_items['type'] == 'earning']
                for _, item in earning_items.iterrows():
                    default_items[item['id']] = st.number_input(f"› {item['name']}", min_value=0.0, step=100.0, value=0.0)
                
                st.write("**減項 (Deductions)**")
                deduction_items = salary_items[salary_items['type'] == 'deduction']
                for _, item in deduction_items.iterrows():
                    default_items[item['id']] = st.number_input(f"› {item['name']}", min_value=0.0, step=100.0, value=0.0)

                st.subheader("步驟二：設定發薪日並產生")
                pay_date = st.date_input("設定發薪日", value=datetime(year, month, monthrange(year, month)[1]))

                submitted = st.form_submit_button("產生本月薪資")
                if submitted:
                    try:
                        with st.spinner("正在為所有在職員工計算薪資..."):
                            emp_count, detail_count = generate_monthly_salaries(conn, year, month, pay_date, default_items)
                        st.success(f"成功產生 {emp_count} 位員工的薪資紀錄，共 {detail_count} 筆明細。")
                        st.info("頁面將在3秒後自動刷新...")
                        st.rerun() # 重新整理頁面以顯示查詢按鈕
                    except Exception as e:
                        st.error(f"產生薪資時發生錯誤：{e}")

    # --- Tab 2: 薪資項目設定 ---
    with tab2:
        st.subheader("薪資項目列表")
        
        try:
            items_df = get_all_salary_items(conn)
            st.dataframe(items_df, use_container_width=True)

            with st.expander("新增或修改薪資項目"):
                # 使用 session state 來儲存正在編輯的項目
                if 'editing_item_id' not in st.session_state:
                    st.session_state.editing_item_id = None
                
                # 顯示選擇框，讓用戶選擇要編輯的項目
                item_list = {"新增項目": None}
                item_list.update({f"{row['name']} (ID: {row['id']})": row['id'] for _, row in items_df.iterrows()})
                
                selected_item_key = st.selectbox("選擇要操作的項目", options=item_list.keys())
                st.session_state.editing_item_id = item_list[selected_item_key]

                # 根據選擇顯示對應的表單
                item_data = {}
                if st.session_state.editing_item_id:
                    # 編輯模式
                    item_data = items_df[items_df['id'] == st.session_state.editing_item_id].iloc[0]
                    form_title = "編輯薪資項目"
                else:
                    # 新增模式
                    form_title = "新增薪資項目"

                with st.form("salary_item_form", clear_on_submit=False):
                    st.write(form_title)
                    name = st.text_input("項目名稱", value=item_data.get('name', ''))
                    type = st.selectbox("類型", ['earning', 'deduction'], index=0 if item_data.get('type', 'earning') == 'earning' else 1)
                    is_active = st.checkbox("啟用中", value=item_data.get('is_active', True))
                    
                    submitted = st.form_submit_button("儲存")
                    if submitted:
                        if not name:
                            st.error("項目名稱為必填！")
                        else:
                            new_data = {'name': name, 'type': type, 'is_active': is_active}
                            try:
                                if st.session_state.editing_item_id:
                                    # 更新
                                    update_salary_item(conn, st.session_state.editing_item_id, new_data)
                                    st.success(f"成功更新項目：{name}")
                                else:
                                    # 新增
                                    add_salary_item(conn, new_data)
                                    st.success(f"成功新增項目：{name}")
                                st.session_state.editing_item_id = None # 清除狀態
                                st.rerun()
                            except Exception as e:
                                st.error(f"操作失敗：{e}")

        except Exception as e:
            st.error(f"讀取薪資項目時發生錯誤：{e}")