import sqlite3

DB_NAME = 'hr_system.db'

def add_employee(name_ch, id_no, entry_date, hr_code=None, gender=None, birth_date=None, phone=None, address=None, dept=None, title=None, resign_date=None, bank_account=None, note=None):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO employee (name_ch, id_no, entry_date, hr_code, gender, birth_date, phone, address, dept, title, resign_date, bank_account, note)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (name_ch, id_no, entry_date, hr_code, gender, birth_date, phone, address, dept, title, resign_date, bank_account, note))
    conn.commit()
    conn.close()
    print("新增員工成功")

def list_employees():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, name_ch, id_no, entry_date, hr_code FROM employee")
    for row in cursor.fetchall():
        print(row)
    conn.close()

def update_employee(emp_id, **kwargs):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    fields = ', '.join([f"{k}=?" for k in kwargs.keys()])
    values = list(kwargs.values()) + [emp_id]
    sql = f"UPDATE employee SET {fields} WHERE id=?"
    cursor.execute(sql, values)
    conn.commit()
    conn.close()
    print("員工資料更新成功")

def delete_employee(emp_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM employee WHERE id=?", (emp_id,))
    conn.commit()
    conn.close()
    print("員工已刪除")

# 以下是簡單的命令列選單（僅供測試）
if __name__ == '__main__':
    print("員工資料管理：")
    print("1. 新增 2. 查詢 3. 修改 4. 刪除")
    opt = input("請輸入功能選項：")
    if opt == "1":
        name_ch = input("姓名：")
        id_no = input("身分證/居留證號：")
        entry_date = input("入職日期（YYYY-MM-DD）：")
        hr_code = input("人資編號（可留空）：") or None
        add_employee(name_ch, id_no, entry_date, hr_code)
    elif opt == "2":
        list_employees()
    elif opt == "3":
        emp_id = int(input("員工ID："))
        field = input("要修改的欄位名（如name_ch）：")
        value = input("新值：")
        update_employee(emp_id, **{field: value})
    elif opt == "4":
        emp_id = int(input("員工ID："))
        delete_employee(emp_id)
    else:
        print("選項錯誤")
