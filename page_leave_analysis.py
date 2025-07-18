import streamlit as st
import pandas as pd
from datetime import datetime
import numpy as np
import traceback
from utils import (
    get_all_employees,
    read_leave_file,
    check_leave_hours,
    find_absence,
    generate_leave_attendance_comparison,
    DEFAULT_GSHEET_URL
)

def show_page(conn):
    """
    顯示請假與異常分析頁面的主函式
    """
    st.header("請假與異常分析")
    
    # 使用 Tabs 將兩個相關功能整合在同一頁
    tab1, tab2, tab3 = st.tabs(["出勤異常查詢 (缺勤比對)", "請假時數核對", "請假與出勤重疊分析"])

    # --- Tab 1: 出勤異常查詢 (缺勤比對) ---
    with tab1:
        st.subheader("比對出勤與請假紀錄，找出缺勤")
        st.info("此功能會比對「打卡紀錄」與「請假紀錄」，找出應出勤但無打卡紀錄的員工。")
        
        method = st.radio("請假紀錄來源", ["Google Sheet 連結", "上傳 CSV 檔"], key='leave_source')
        leave_df = None
        
        if method == "Google Sheet 連結":
            url = st.text_input("Google Sheet CSV 連結", value=DEFAULT_GSHEET_URL)
            if url:
                try:
                    leave_df = read_leave_file(url)
                    st.success("成功讀取 Google Sheet 請假紀錄！")
                except Exception as e:
                    st.error(f"讀取 Google Sheet 時發生錯誤: {e}")
        else:
            uploaded_leave_file = st.file_uploader("上傳請假 CSV 檔案", type=['csv'], key='leave_file_uploader')
            if uploaded_leave_file:
                try:
                    leave_df = read_leave_file(uploaded_leave_file)
                    st.success("成功讀取上傳的 CSV 請假紀錄！")
                except Exception as e:
                    st.error(f"讀取 CSV 檔案時發生錯誤: {e}")

        st.write("---")
        today = datetime.now()
        c1, c2 = st.columns(2)
        year = c1.number_input("查詢年份", min_value=2020, max_value=today.year + 1, value=today.year)
        month = c2.number_input("查詢月份", min_value=1, max_value=12, value=today.month)
        
        mode = st.radio("報表模式", ["缺勤明細", "缺勤彙總"], horizontal=True)
        
        if st.button("產生缺勤報告", key="generate_absence_report"):
            if leave_df is None:
                st.warning("請先提供並成功讀取請假紀錄！")
            else:
                try:
                    attendance_df = pd.read_sql_query("SELECT * FROM attendance", conn)
                    emp_df = get_all_employees(conn)
                    
                    if attendance_df.empty:
                        st.warning("資料庫中沒有任何出勤紀錄可供比對。")
                    else:
                        with st.spinner("正在比對資料..."):
                            count_mode = (mode == "缺勤彙總")
                            df_absent = find_absence(attendance_df, leave_df, emp_df, year, month, count_mode=count_mode)
                        
                        st.write("---")
                        st.subheader("缺勤報告結果")
                        st.dataframe(df_absent)
                        
                        if not df_absent.empty:
                            fname = f"absence_report_{year}-{month:02d}.csv"
                            st.download_button(
                                "下載報告 CSV",
                                df_absent.to_csv(index=False).encode("utf-8-sig"),
                                file_name=fname
                            )
                except Exception as ex:
                    st.error(f"查詢失敗: {ex}")

    # --- Tab 2: 請假時數核對 ---
    with tab2:
        st.subheader("核對請假單時數")
        st.info("此功能會讀取請假紀錄，並根據您選擇的月份和人員進行篩選，同時標示出系統核算與原始時數不符的異常紀錄。")
        
        leave_source = st.radio("請假紀錄來源", ["Google Sheet 連結", "上傳 CSV 檔"], key='leave_source_check', horizontal=True)
        
        source_input = None
        if leave_source == "Google Sheet 連結":
            source_input = st.text_input("Google Sheet CSV 連結", value=DEFAULT_GSHEET_URL, key="url_check")
        else:
            source_input = st.file_uploader("上傳請假 CSV 檔案", type=['csv'], key='leave_file_uploader_check')

        st.write("---")
        st.markdown("#### 篩選條件")
        
        try:
            all_employees_df = get_all_employees(conn)
            emp_list = all_employees_df['name_ch'].tolist()
            selected_employees = st.multiselect("選擇員工 (可多選，留空則代表全部員工)", options=emp_list)
        except Exception as e:
            st.error(f"讀取員工列表失敗: {e}")
            selected_employees = []

        c1, c2 = st.columns(2)
        today = datetime.now()
        selected_year = c1.number_input("年份", min_value=2020, max_value=today.year + 1, value=today.year, key="check_year")
        selected_month = c2.number_input("月份", min_value=1, max_value=12, value=today.month, key="check_month")
        
        if st.button("讀取並核對時數", key="check_hours_button"):
            # 按下按鈕時，執行計算並將結果存入 session_state
            if not source_input:
                st.warning("請提供請假紀錄的來源 (連結或檔案)。")
                st.session_state['leave_check_results'] = None # 清空舊結果
            else:
                try:
                    with st.spinner("正在讀取、計算並篩選資料..."):
                        leave_df = read_leave_file(source_input)
                        check_df = check_leave_hours(leave_df)
                    
                    if not check_df.empty:
                        # 進行篩選
                        filtered_df = check_df[
                            (check_df['Start Date'].dt.year == selected_year) &
                            (check_df['Start Date'].dt.month == selected_month)
                        ].copy()

                        if selected_employees:
                            filtered_df = filtered_df[filtered_df['Employee Name'].isin(selected_employees)]

                        # 標示異常
                        filtered_df['異常註記'] = np.where(
                            ~np.isclose(filtered_df['Duration'].astype(float), filtered_df['核算時數'].astype(float)), 
                            "時數不符", ""
                        )
                        # *** 關鍵修正：將處理好的結果存入 session_state ***
                        st.session_state['leave_check_results'] = filtered_df
                    else:
                        st.session_state['leave_check_results'] = pd.DataFrame() # 存入空的DataFrame
                except Exception as e:
                    st.error(f"處理失敗: {e}")
                    st.error(traceback.format_exc())
                    st.session_state['leave_check_results'] = None # 清空結果

        # --- *** 關鍵修正：將顯示邏輯移到按鈕區塊之外 *** ---
        # 每次頁面刷新時，都檢查 session_state 中是否有結果可顯示
        if 'leave_check_results' in st.session_state and st.session_state['leave_check_results'] is not None:
            results_df = st.session_state['leave_check_results']
            
            st.write("---")
            st.subheader("核對結果")

            show_only_anomalies = st.checkbox("僅顯示異常紀錄")
            
            if show_only_anomalies:
                display_df = results_df[results_df['異常註記'] == "時數不符"]
            else:
                display_df = results_df

            if display_df.empty:
                st.info("在當前篩選條件下，沒有找到可顯示的紀錄。")
            else:
                st.dataframe(display_df[['Employee Name','Start Date','End Date','Type of Leave','Duration','核算時數', '異常註記']])
                
                fname = f"leave_hours_check_{selected_year}-{selected_month:02d}.csv"
                st.download_button(
                    "下載當前檢視的報表",
                    display_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name=fname
                )

    # --- Tab 3: 請假與出勤交叉比對 [功能升級] ---
    with tab3:
        st.subheader("比對請假紀錄與實際打卡狀況")
        st.info("此功能會比對指定月份的請假單與打卡紀錄，並標示出各種狀態，您可以篩選出需要關注的異常狀況。")
        
        method_conflict = st.radio("請假紀錄來源", ["Google Sheet 連結", "上傳 CSV 檔"], key='conflict_source', horizontal=True)
        leave_df_conflict_source = None
        
        if method_conflict == "Google Sheet 連結":
            leave_df_conflict_source = st.text_input("Google Sheet CSV 連結", value=DEFAULT_GSHEET_URL, key="url_conflict")
        else:
            leave_df_conflict_source = st.file_uploader("上傳請假 CSV 檔案", type=['csv'], key='leave_file_uploader_conflict')

        st.write("---")
        st.markdown("#### 請選擇比對月份")
        c1, c2 = st.columns(2)
        today = datetime.now()
        year = c1.number_input("年份", min_value=2020, max_value=today.year + 1, value=today.year, key="conflict_year")
        month = c2.number_input("月份", min_value=1, max_value=12, value=today.month, key="conflict_month")
        
        if st.button("開始交叉比對", key="conflict_button"):
            if not leave_df_conflict_source:
                st.warning("請先提供並成功讀取請假紀錄！")
                st.session_state['comparison_results'] = None
            else:
                try:
                    with st.spinner("正在讀取資料庫並進行比對..."):
                        leave_df = read_leave_file(leave_df_conflict_source)
                        attendance_df = pd.read_sql_query("SELECT * FROM attendance", conn)
                        emp_df = get_all_employees(conn)
                        
                        comparison_df = generate_leave_attendance_comparison(leave_df, attendance_df, emp_df, year, month)
                        st.session_state['comparison_results'] = comparison_df
                except Exception as e:
                    st.error(f"比對過程中發生錯誤: {e}")
                    st.error(traceback.format_exc())
                    st.session_state['comparison_results'] = None

        if 'comparison_results' in st.session_state and st.session_state['comparison_results'] is not None:
            results_df = st.session_state['comparison_results']
            
            st.write("---")
            st.subheader("交叉比對結果")

            if results_df.empty:
                st.success(f"在 {year} 年 {month} 月中，未發現任何請假或出勤紀錄可供比對。")
            else:
                # --- *** 這是關鍵的修正點 *** ---
                # 1. 修改勾選框的標題文字
                show_only_anomalies = st.checkbox("僅顯示異常紀錄", key="conflict_anomalies_check")
                st.caption("異常紀錄包含：「請假期間有打卡」和「無出勤且無請假紀錄」。")
                
                if show_only_anomalies:
                    # 2. 修改篩選邏輯，使用 isin 來包含兩種異常狀態
                    anomalous_statuses = ['異常：請假期間有打卡', '無出勤且無請假紀錄']
                    display_df = results_df[results_df['狀態註記'].isin(anomalous_statuses)]
                else:
                    display_df = results_df
                
                if display_df.empty:
                    st.info("在當前條件下無符合的紀錄。")
                else:
                    st.dataframe(display_df, use_container_width=True)
                    
                    fname = f"leave_attendance_comparison_{year}-{month:02d}.csv"
                    st.download_button(
                        "下載當前檢視的報告CSV",
                        display_df.to_csv(index=False).encode("utf-8-sig"),
                        file_name=fname
                    )