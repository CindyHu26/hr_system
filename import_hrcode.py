import pandas as pd
import sqlite3

DB_NAME = 'hr_system.db'
EMPLOYEE_FILE = '11406.xlsx'

# 讀入Excel
df = pd.read_excel(EMPLOYEE_FILE, dtype=str)

# 自動抓欄位
name_col = None
phone_col = None
for c in df.columns:
    if '名' in c or 'name' in c.lower():
        name_col = c
    if '電話' in c or 'phone' in c.lower():
        phone_col = c

if not name_col or not phone_col:
    raise Exception("找不到姓名或電話欄位，請檢查Excel！")

conn = sqlite3.connect(DB_NAME)
updated = 0
for _, row in df.iterrows():
    name_ch = str(row[name_col]).strip()
    phone = str(row[phone_col]).strip()
    # 若以數字開頭且長度為9~10且不是09開頭，補0
    if phone and not phone.startswith('0'):
        if len(phone) in [8,9,10]:
            phone = '0' + phone
    # 避免已經有0的再重複
    # 只更新非空資料
    if not name_ch or not phone:
        continue
    cur = conn.execute("UPDATE employee SET phone=? WHERE name_ch=?", (phone, name_ch))
    if cur.rowcount > 0:
        updated += 1
conn.commit()
conn.close()
print(f"批次更新完成，共更新 {updated} 筆員工的電話。")

# 檢查有無找不到對應員工
no_match = df[~df[name_col].isin(pd.read_sql("SELECT name_ch FROM employee", sqlite3.connect(DB_NAME))["name_ch"])]
if not no_match.empty:
    print("下列員工在資料庫找不到，請人工核查：")
    print(no_match)
