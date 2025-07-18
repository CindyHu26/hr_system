import pandas as pd
import sqlite3

DB_NAME = 'hr_system.db'

employee_file = 'employee.csv'
company_file = 'company.csv'
history_file = 'employee_company_history.csv'

def import_company(conn, file):
    df = pd.read_csv(file, encoding='utf-8-sig', dtype=str)
    for _, row in df.iterrows():
        c = conn.execute("SELECT id FROM company WHERE name=?", (row['name'],)).fetchone()
        if not c:
            conn.execute(
                "INSERT INTO company (name, uniform_no) VALUES (?, ?)",
                (row['name'], row.get('uniform_no'))
            )
    conn.commit()
    print("公司資料匯入完成")

def import_employee(conn, file):
    df = pd.read_csv(file, encoding='utf-8-sig', dtype=str)
    for _, row in df.iterrows():
        entry_date = row.get('entry_date') if row.get('entry_date') and str(row.get('entry_date')).strip() != '' else None
        if not entry_date:
            print(f"警告：{row.get('name_ch')}（身分證：{row.get('id_no')}）缺少入職日，將以空值匯入")
        c = conn.execute("SELECT id FROM employee WHERE id_no=?", (row['id_no'],)).fetchone()
        if not c:
            conn.execute(
                "INSERT INTO employee (name_ch, id_no, entry_date, birth_date, address) VALUES (?, ?, ?, ?, ?)",
                (row.get('name_ch'), row.get('id_no'), entry_date, row.get('birth_date'), row.get('address'))
            )
    conn.commit()
    print("員工資料匯入完成")

def import_history(conn, file):
    df = pd.read_csv(file, encoding='utf-8-sig', dtype=str)
    for _, row in df.iterrows():
        emp = conn.execute("SELECT id FROM employee WHERE id_no=?", (row['emp_id_no'],)).fetchone()
        comp = conn.execute("SELECT id FROM company WHERE name=?", (row['company_name'],)).fetchone()
        if emp and comp:
            conn.execute(
                "INSERT INTO employee_company_history (employee_id, company_id, start_date, note) VALUES (?, ?, ?, ?)",
                (emp[0], comp[0], row.get('start_date'), row.get('note'))
            )
        else:
            print(f"未找到對應，略過：{row['emp_name']} / {row['company_name']}")
    conn.commit()
    print("加保異動資料匯入完成")

if __name__ == '__main__':
    conn = sqlite3.connect(DB_NAME)
    import_company(conn, company_file)
    import_employee(conn, employee_file)
    import_history(conn, history_file)
    conn.close()
    print("全部匯入完成")
