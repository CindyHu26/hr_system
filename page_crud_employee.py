import streamlit as st
import pandas as pd
from datetime import datetime
import sqlite3
from utils import (
    get_all_employees,
    get_employee_by_id,
    add_employee,
    update_employee,
    delete_employee,
    show_salary_base_history_tab,
    EMPLOYEE_COLUMNS_MAP
)

def show_page(conn):
    """
    顯示員工資料管理頁面 (CRUD) 的主函式
    """
    st.header("員工資料管理")
    tabs = st.tabs(["員工資料", "底薪／眷屬異動"])
    with tabs[0]:
        # --- 顯示與篩選員工 ---
        st.subheader("員工列表")
        try:
            # 從 utils 取得所有員工資料，並將欄位重命名為中文
            all_emp_df_raw = get_all_employees(conn)
            all_emp_df_display = all_emp_df_raw.rename(columns=EMPLOYEE_COLUMNS_MAP)
            st.dataframe(all_emp_df_display)
        except Exception as e:
            st.error(f"無法讀取員工資料: {e}")
            # 如果無法讀取員工資料，後續操作無意義，直接返回
            return

        st.subheader("資料操作")
        crud_option = st.selectbox("選擇操作", ["新增 (Create)", "修改 (Update) / 刪除 (Delete)"])

        # --- 新增員工 ---
        if crud_option == "新增 (Create)":
            with st.form("add_employee_form", clear_on_submit=True):
                st.write("請填寫新員工資料：")
                c1, c2, c3 = st.columns(3)
                
                # 使用 st.date_input 讓使用者用日曆選擇日期
                name_ch_add = c1.text_input(EMPLOYEE_COLUMNS_MAP['name_ch'], key="add_name")
                hr_code_add = c2.text_input(EMPLOYEE_COLUMNS_MAP['hr_code'], key="add_hr_code")
                id_no_add = c3.text_input(EMPLOYEE_COLUMNS_MAP['id_no'], key="add_id_no")
                dept_add = c1.text_input(EMPLOYEE_COLUMNS_MAP['dept'], key="add_dept")
                title_add = c2.text_input(EMPLOYEE_COLUMNS_MAP['title'], key="add_title")
                gender_add = c3.selectbox(EMPLOYEE_COLUMNS_MAP['gender'], [None, "男", "女"], key="add_gender")
                entry_date_add = c1.date_input(EMPLOYEE_COLUMNS_MAP['entry_date'], value=None, key="add_entry_date")
                birth_date_add = c2.date_input(EMPLOYEE_COLUMNS_MAP['birth_date'], value=None, key="add_birth_date")
                resign_date_add = c3.date_input(EMPLOYEE_COLUMNS_MAP['resign_date'], value=None, key="add_resign_date")
                phone_add = c1.text_input(EMPLOYEE_COLUMNS_MAP['phone'], key="add_phone")
                address_add = c2.text_input(EMPLOYEE_COLUMNS_MAP['address'], key="add_address")
                bank_account_add = c3.text_input(EMPLOYEE_COLUMNS_MAP['bank_account'], key="add_bank")
                note_add = st.text_area(EMPLOYEE_COLUMNS_MAP['note'], key="add_note")

                submitted = st.form_submit_button("新增員工")
                if submitted:
                    # 收集表單資料
                    new_data = {
                        'name_ch': name_ch_add, 'hr_code': hr_code_add, 'id_no': id_no_add,
                        'dept': dept_add, 'title': title_add, 'gender': gender_add,
                        'entry_date': entry_date_add, 'birth_date': birth_date_add, 'resign_date': resign_date_add,
                        'phone': phone_add, 'address': address_add, 'bank_account': bank_account_add,
                        'note': note_add
                    }

                    # 淨化資料，將空值或空字串統一為 None
                    for key, value in new_data.items():
                        if not value:
                            new_data[key] = None
                    
                    if not new_data['name_ch'] or not new_data['hr_code']:
                        st.error("姓名與員工編號為必填欄位！")
                    else:
                        try:
                            add_employee(conn, new_data)
                            st.success(f"成功新增員工：{new_data['name_ch']}")
                            st.rerun()
                        except sqlite3.IntegrityError as e:
                            st.error(f"新增失敗：員工編號或身份證號可能已存在。 {e}")
                        except Exception as e:
                            st.error(f"發生未知錯誤：{e}")

        # --- 修改/刪除員工 ---
        elif crud_option == "修改 (Update) / 刪除 (Delete)":
            st.write("請先從下方選擇一位員工進行操作：")
            if not all_emp_df_raw.empty:
                # 建立選擇列表
                options_df = all_emp_df_raw[['id', 'name_ch', 'hr_code']].copy()
                options_df['display'] = options_df['name_ch'] + " (" + options_df['hr_code'].astype(str) + ")"
                selected_display = st.selectbox("選擇員工", options=options_df['display'])
                
                if selected_display:
                    selected_id = int(options_df[options_df['display'] == selected_display]['id'].iloc[0])
                    selected_employee = get_employee_by_id(conn, selected_id)

                    if selected_employee is not None:
                        st.write(f"### 正在編輯: {selected_employee['name_ch']}")
                        
                        # 輔助函式：將資料庫中的日期字串安全地轉換為 date 物件
                        def to_date(date_string):
                            if date_string:
                                try:
                                    return datetime.strptime(date_string, '%Y-%m-%d').date()
                                except (ValueError, TypeError):
                                    return None
                            return None

                        with st.form("update_employee_form"):
                            c1, c2, c3 = st.columns(3)
                            name_ch_input = c1.text_input(EMPLOYEE_COLUMNS_MAP['name_ch'], value=selected_employee.get('name_ch', ''))
                            hr_code_input = c2.text_input(EMPLOYEE_COLUMNS_MAP['hr_code'], value=selected_employee.get('hr_code', ''))
                            id_no_input = c3.text_input(EMPLOYEE_COLUMNS_MAP['id_no'], value=selected_employee.get('id_no', ''))
                            dept_input = c1.text_input(EMPLOYEE_COLUMNS_MAP['dept'], value=selected_employee.get('dept', '') or '')
                            title_input = c2.text_input(EMPLOYEE_COLUMNS_MAP['title'], value=selected_employee.get('title', '') or '')
                            
                            gender_options = [None, "男", "女"]
                            current_gender = selected_employee.get('gender')
                            gender_index = gender_options.index(current_gender) if current_gender in gender_options else 0
                            gender_input = c3.selectbox(EMPLOYEE_COLUMNS_MAP['gender'], gender_options, index=gender_index)
                            
                            entry_date_input = c1.date_input(EMPLOYEE_COLUMNS_MAP['entry_date'], value=to_date(selected_employee.get('entry_date')))
                            birth_date_input = c2.date_input(EMPLOYEE_COLUMNS_MAP['birth_date'], value=to_date(selected_employee.get('birth_date')))
                            resign_date_input = c3.date_input(EMPLOYEE_COLUMNS_MAP['resign_date'], value=to_date(selected_employee.get('resign_date')))
                            
                            phone_input = c1.text_input(EMPLOYEE_COLUMNS_MAP['phone'], value=selected_employee.get('phone', '') or '')
                            address_input = c2.text_input(EMPLOYEE_COLUMNS_MAP['address'], value=selected_employee.get('address', '') or '')
                            bank_account_input = c3.text_input(EMPLOYEE_COLUMNS_MAP['bank_account'], value=selected_employee.get('bank_account', '') or '')
                            note_input = st.text_area(EMPLOYEE_COLUMNS_MAP['note'], value=selected_employee.get('note', '') or '')
                            
                            update_button = st.form_submit_button("儲存變更")
                            
                            if update_button:
                                # 淨化並收集表單資料
                                updated_data = {
                                    'name_ch': name_ch_input or None, 'hr_code': hr_code_input or None, 'id_no': id_no_input or None,
                                    'dept': dept_input or None, 'title': title_input or None, 'gender': gender_input,
                                    'entry_date': entry_date_input, 'birth_date': birth_date_input, 'resign_date': resign_date_input,
                                    'phone': phone_input or None, 'address': address_input or None, 'bank_account': bank_account_input or None,
                                    'note': note_input or None,
                                }
                                
                                try:
                                    result = update_employee(conn, selected_id, updated_data)
                                    if result > 0:
                                        st.success(f"成功更新員工 {updated_data['name_ch']} 的資料！")
                                        st.rerun()
                                    else:
                                        st.error("更新失敗，影響 0 筆資料。")
                                except Exception as e:
                                    st.error(f"更新時發生錯誤：{e}")

                        if st.button("🔴 刪除這位員工", key=f"delete_{selected_id}"):
                            st.warning(f"您確定要永久刪除 **{selected_employee['name_ch']}** 嗎？此操作無法復原！")
                            if st.button("我確定，請刪除", key=f"confirm_delete_{selected_id}"):
                                try:
                                    delete_employee(conn, selected_id)
                                    st.success(f"已成功刪除員工 {selected_employee['name_ch']}。")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"刪除失敗：{e} (該員工可能仍有關聯的出勤紀錄)")
    
    with tabs[1]:
        show_salary_base_history_tab(conn)