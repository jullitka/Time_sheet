import sqlite3


def create_db():
    """Создает базу данных сотрудников"""
    conn = sqlite3.connect('employees.db')
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS employees(
                employee_id INTEGER PRIMARY KEY AUTOINCREMENT,
                last_name TEXT NOT NULL,
                first_name TEXT NOT NULL,
                patronymic TEXT NOT NULL,
                time_sheet_number TEXT,
                position TEXT NOT NULL,
                part_time REAL NOT NULL,
                internal_combine TEXT,
                employment_date TEXT,
                dismissal_date TEXT,
                department TEXT
                );
                """)

    cur.execute("""CREATE TABLE IF NOT EXISTS schedule(
                shedule_id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER,
                Monday INTEGER,
                Tuesdau INTEGER,
                Wednesday INTEGER,
                Thursday INTEGER,
                Friday INTEGER,
                FOREIGN KEY (employee_id) REFERENCES employees (employee_id)
                );
                """)
    conn.commit()

# create_db()

# используется
def add_employee_to_db(last_name, first_name, patronymic,
                       time_sheet_number, position, part_time,
                       internal_combine, employment_date,
                       dismissal_date, department):
    """Добавляет нового сотрудника в БД."""
    conn = sqlite3.connect('employees.db')
    cur = conn.cursor()
    cur.execute("""INSERT INTO employees
                (last_name, first_name, patronymic,
                 time_sheet_number, position, part_time,
                 internal_combine, employment_date,
                 dismissal_date, department)
                VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?);""",
                (last_name, first_name, patronymic,
                 time_sheet_number, position, part_time,
                 internal_combine, employment_date,
                 dismissal_date, department)
                )
    conn.commit()

# используется
def del_employee_from_db(last_name, first_name, patronymic,
                         time_sheet_number, part_time):
    """Удаляет сотрудника из БД."""
    conn = sqlite3.connect('employees.db')
    cur = conn.cursor()
    cur.execute("""DELETE FROM employees
                WHERE last_name=? AND first_name=?
                AND patronymic=? AND part_time=?
                AND time_sheet_number=?;""",
                (last_name, first_name, patronymic,
                 part_time, time_sheet_number))
    conn.commit()


# используется
def search_last_name(last_name=None):
    """Поиск сотрудников по совпадению фамилии.
    Возвращает список сотрудников.
    Если last_name не указан, то возвращает список всех
    сотрудников"""
    conn = sqlite3.connect('employees.db')
    cur = conn.cursor()
    if last_name:
        cur.execute("""SELECT last_name, first_name, patronymic,
                    time_sheet_number, position, part_time
                    FROM employees
                    WHERE last_name=?;""",
                    (last_name,))
    else:
        cur.execute("""SELECT last_name, first_name, patronymic,
                    time_sheet_number, position, part_time, internal_combine
                    FROM employees;""")
        
    listemployees = cur.fetchall()
    conn.commit()
    return listemployees


# используется
def search_employee_in_db(last_name, first_name, patronymic,
                          time_sheet_number, position, part_time,
                          internal_combine, department):
    """Ищет сотрудника в базе данных.
    Возвращает сотрудника"""
    conn = sqlite3.connect('employees.db')
    cur = conn.cursor()
    cur.execute("""SELECT * FROM employees
                WHERE last_name=? AND first_name=? 
                AND patronymic=? AND internal_combine=?
                AND time_sheet_number=? AND position=?
                AND part_time=? AND department=?""",
                (last_name, first_name, patronymic,
                 internal_combine, time_sheet_number,
                 position, part_time, department))
    employee = cur.fetchone()
    return employee

