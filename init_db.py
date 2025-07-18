import sqlite3

# 資料庫檔案名稱
DB_NAME = 'hr_system.db'

def create_tables(conn):
    cursor = conn.cursor()

    # 員工資料表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS employee (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name_ch TEXT NOT NULL,
        id_no TEXT NOT NULL,
        entry_date DATE,              -- 允許NULL
        hr_code TEXT,
        gender TEXT,
        birth_date DATE,
        phone TEXT,
        address TEXT,
        dept TEXT,
        title TEXT,
        resign_date DATE,
        bank_account TEXT,
        note TEXT
    )
    """)

    # 公司（加保單位）表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS company (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        uniform_no TEXT,
        address TEXT,
        owner TEXT,
        ins_code TEXT,
        note TEXT
    )
    """)

    # 員工加保異動紀錄表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS employee_company_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER NOT NULL,
        company_id INTEGER NOT NULL,
        start_date DATE NOT NULL,
        end_date DATE,
        note TEXT,
        FOREIGN KEY(employee_id) REFERENCES employee(id),
        FOREIGN KEY(company_id) REFERENCES company(id)
    )
    """)

    # 出勤紀錄表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER NOT NULL,
        date DATE NOT NULL,
        checkin_time TIME,
        checkout_time TIME,
        late_minutes INTEGER DEFAULT 0,
        early_leave_minutes INTEGER DEFAULT 0,
        absent_minutes INTEGER DEFAULT 0,
        overtime1_minutes INTEGER DEFAULT 0,
        overtime2_minutes INTEGER DEFAULT 0,
        overtime3_minutes INTEGER DEFAULT 0,
        note TEXT,
        source_file TEXT,
        FOREIGN KEY(employee_id) REFERENCES employee(id),
        UNIQUE(employee_id, date) -- 同一個員工同一天只能有一筆紀錄
    )
    """)

    # 請假紀錄表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS leave_record (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER NOT NULL,
        request_id TEXT,
        leave_type TEXT NOT NULL,
        start_date DATE NOT NULL,
        end_date DATE NOT NULL,
        duration FLOAT,
        reason TEXT,
        status TEXT,
        approver TEXT,
        submit_date DATE,
        note TEXT,
        FOREIGN KEY(employee_id) REFERENCES employee(id)
    )
    """)

    # 薪資項目定義表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS salary_item (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        type TEXT NOT NULL,  -- earning / deduction
        is_active BOOLEAN NOT NULL DEFAULT 1
    )
    """)

    # 薪資紀錄主表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS salary (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER NOT NULL,
        year INTEGER NOT NULL,
        month INTEGER NOT NULL,
        pay_date DATE,
        note TEXT,
        FOREIGN KEY(employee_id) REFERENCES employee(id)
    )
    """)

    # 薪資明細表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS salary_detail (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        salary_id INTEGER NOT NULL,
        salary_item_id INTEGER NOT NULL,
        amount FLOAT NOT NULL,
        FOREIGN KEY(salary_id) REFERENCES salary(id),
        FOREIGN KEY(salary_item_id) REFERENCES salary_item(id)
    )
    """)

    conn.commit()
    print("資料表建立完成！")

if __name__ == '__main__':
    conn = sqlite3.connect(DB_NAME)
    create_tables(conn)
    conn.close()
    print(f'資料庫 {DB_NAME} 初始化完成。')
