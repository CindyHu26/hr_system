import streamlit as st
import pandas as pd
from utils import get_all_employees

def employee_selector(conn, key_prefix="", pre_selected_ids=None):
    """
    一個可重複使用的員工選擇器元件，具備部門篩選和預選功能。
    返回選定的員工 ID 列表。
    """
    st.markdown("##### 選擇員工")
    
    if pre_selected_ids is None:
        pre_selected_ids = []
    
    try:
        emp_df = get_all_employees(conn)[['id', 'name_ch', 'dept', 'title']]
        emp_df['display'] = emp_df['name_ch'] + " (" + emp_df['dept'].fillna('未分配') + " - " + emp_df['title'].fillna('無職稱') + ")"
        
        # 1. 建立部門篩選器
        valid_depts = sorted([dept for dept in emp_df['dept'].unique() if pd.notna(dept)])
        all_depts = ['所有部門'] + valid_depts
        
        selected_dept = st.selectbox(
            "依部門篩選", 
            options=all_depts, 
            key=f"{key_prefix}_dept_filter"
        )

        # 2. 根據篩選結果，決定要顯示的員工列表
        if selected_dept == '所有部門':
            filtered_emp_df = emp_df
        else:
            filtered_emp_df = emp_df[emp_df['dept'] == selected_dept]
        
        emp_options = dict(zip(filtered_emp_df['display'], filtered_emp_df['id']))

        # 3. **核心修正：找出預設選項**
        # 建立一個反向對應字典 { id: '姓名 (部門 - 職稱)' }
        id_to_display_map = {v: k for k, v in emp_options.items()}
        default_selections = [id_to_display_map[emp_id] for emp_id in pre_selected_ids if emp_id in id_to_display_map]

        # 4. 建立員工複選框，並傳入預設值
        selected_emp_displays = st.multiselect(
            "員工列表 (可複選)",
            options=list(emp_options.keys()),
            default=default_selections, # <-- 將預選名單設定於此
            key=f"{key_prefix}_multiselect"
        )
        
        # 5. 返回選定的員工 ID 列表
        selected_employee_ids = [emp_options[display] for display in selected_emp_displays]
        return selected_employee_ids

    except Exception as e:
        st.error(f"載入員工選擇器時發生錯誤: {e}")
        return []