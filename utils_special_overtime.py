# utils_special_overtime.py
import pandas as pd
from datetime import datetime, time

def get_special_attendance(conn, year, month):
    query = """
    SELECT sa.id, e.name_ch as '員工姓名', sa.date as '日期', sa.checkin_time as '上班時間', sa.checkout_time as '下班時間', sa.note as '備註'
    FROM special_attendance sa
    JOIN employee e ON sa.employee_id = e.id
    WHERE STRFTIME('%Y-%m', sa.date) = ?
    ORDER BY sa.date, e.name_ch
    """
    month_str = f"{year}-{month:02d}"
    return pd.read_sql_query(query, conn, params=(month_str,))

def add_special_attendance(conn, data):
    cursor = conn.cursor()
    sql = "INSERT INTO special_attendance (employee_id, date, checkin_time, checkout_time, note) VALUES (?, ?, ?, ?, ?)"
    cursor.execute(sql, (data['employee_id'], data['date'], data['checkin_time'], data['checkout_time'], data['note']))
    conn.commit()

def delete_special_attendance(conn, record_id):
    cursor = conn.cursor()
    cursor.execute("DELETE FROM special_attendance WHERE id = ?", (record_id,))
    conn.commit()
    return cursor.rowcount

def calculate_special_overtime_pay(conn, employee_id, year, month, hourly_rate):
    month_str = f"{year}-{month:02d}"
    query = "SELECT checkin_time, checkout_time FROM special_attendance WHERE employee_id = ? AND STRFTIME('%Y-%m', date) = ?"
    records = conn.cursor().execute(query, (employee_id, month_str)).fetchall()

    total_pay = 0
    for checkin_str, checkout_str in records:
        checkin_t = datetime.strptime(checkin_str, '%H:%M:%S').time()
        checkout_t = datetime.strptime(checkout_str, '%H:%M:%S').time()

        duration_hours = (datetime.combine(datetime.min, checkout_t) - datetime.combine(datetime.min, checkin_t)).total_seconds() / 3600

        if duration_hours <= 2:
            pay = duration_hours * hourly_rate * 1.34
        else:
            pay = (2 * hourly_rate * 1.34) + ((duration_hours - 2) * hourly_rate * 1.67)
        total_pay += pay

    return int(round(total_pay))