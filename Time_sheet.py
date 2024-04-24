from tkinter import (BooleanVar, Menu, messagebox,
                     Radiobutton, Tk)
from tkinter import filedialog, ttk
from tkcalendar import DateEntry

from add_data import add_data_from_db
from db import (add_employee_to_db, del_employee_from_db,
                search_employee_in_db, search_last_name)
from classes import BaseForm
from information import COMBINE, DEPARTMENTS, PART_TIME, POSITIONS
from main_program import main_program
from utils import check_employee


class App(Tk):
    def __init__(self):
        super().__init__()
        
        self.title('Time_sheet')
        self.geometry('800x390')
        self['background'] = '#EBEBEB'
        self.conf = {'padx': (10, 30), 'pady': 10}
        self.bolt_font = 'Times New Roman 12'
        self.hellow()

        # Добавление меню
        menu = Menu(self)
        employees_menu = Menu(menu, tearoff=0)
        timesheet_menu = Menu(menu, tearoff=0)

        timesheet_menu.add_command(
            label="Заполнить табель", command=self.put_frame_time_sheet
        )
        # timesheet_menu.add_command(label="Последний табель")
        employees_menu.add_command(
            label="Добавить сотрудника",
            command=self.put_frame_add_employee
        )
        # employees_menu.add_command(
        #     label="Добавить список сотрудников",
        #     command=self.put_frame_add_list_employees
        # )
        employees_menu.add_command(
            label="Удалить сотрудника",
            command=self.put_frame_del_employee
        )
       # employees_menu.add_command(label="Удалить всех сотрудников")
       # employees_menu.add_command(label="Изменить данные сотрудника")
       # employees_menu.add_command(label="Посмотреть список сотрудников")

        menu.add_cascade(label="Табель", menu=timesheet_menu)
        menu.add_cascade(label="Сотрудники", menu=employees_menu)
        # menu.add_command(label="О программе")
        menu.add_command(label="Выйти", command=self.destroy)
        self.config(menu=menu)

    # Форма, которая открывается при открытии программы
    def hellow(self):
        self.form_time_sheet = HellowForm(self).grid(
            row=0, column=0, sticky='nswe'
        )

    # Форма добавления списка сотрудников
    def put_frame_add_list_employees(self):
        self.form_time_sheet = FormAddListEmployees(self).grid(
            row=0, column=0, sticky='nswe'
        )

    # Форма добавления сотрудника
    def put_frame_add_employee(self):
        self.form_time_sheet = FormAddEmployee(self).grid(
            row=0, column=0, sticky='nswe'
        )

    # Форма заполнения табеля
    def put_frame_time_sheet(self):
        self.form_time_sheet = FormTimeSheet(self).grid(
            row=0, column=0, sticky='nswe'
        )

    # Форма удаления сотрудника
    def put_frame_del_employee(self):
        self.form_del_employee = FormDelEmployee(self).grid(
            row=0, column=0, sticky='nswe'
        )


class HellowForm(BaseForm):
    """Форма приветствия"""
    def put_widgets(self):
        self.TEXT = "Добро пожаловать в программу заполнения табеля!"
        self.text = ttk.Label(self, text=self.TEXT)
        self.text.grid(column=0, row=0, sticky='n', pady=(20,10), padx=20)

        # Выбор отдела
        #self.department = ttk.Label(self, text="Выберите отдел")
        #self.department.grid(column=0, row=1, sticky='w', cnf=self.master.conf)

        #self.combo_department = ttk.Combobox(self, values=DEPARTMENTS)
        #self.combo_department.grid(
        #    column=1, row=1, sticky='e', cnf=self.master.conf
        #)


class FormAddListEmployees(BaseForm):
    """Форма добавления списка сотрудников"""
    def put_widgets(self):
        # Выбор файла списка сотрудников
        self.list_employees = ttk.Label(
            self,
            text="Выберите файл, содержащий список сотрудников*"
        )
        self.list_employees.grid(column=0, row=2,
                                 sticky='w', cnf=self.master.conf)

        self.btn_list_employees = ttk.Button(self, text="Выбрать файл",
                                             command=self.choose_file)
        self.btn_list_employees.grid(column=1, row=2,
                                     sticky='e', cnf=self.master.conf)

        # Кнопка запуска добавления сотрудников
        self.btn = ttk.Button(
            self, text="Добавить сотрудников", command=self.add_employees
        )
        self.btn.grid(
            column=0, row=10, columnspan=2, sticky='n', padx=10, pady=10
        )

    def choose_file(self):
        # Исполняется при нажатии кнопки выбора файла со списком сотрудников
        filetypes = (("Таблица", "*.xls, *.xlsx"),)
        filename = filedialog.askopenfilename(
            title="Открыть файл", initialdir="/",
            filetypes=filetypes
        )
        if filename:
            self.file_name = ttk.Label(
                self,
                text=filename
            )
            self.file_name.grid(column=1, row=3,
                                sticky='e', cnf=self.master.conf)
        self.file_list_employees = filename

    def add_employees(self):
        """Добавляет сотрудников в базу данных"""
        add_data_from_db(self.file_list_employees)
        messagebox.showinfo("Сообщение",  "Сотрудники успешно добавлены!")


class FormAddEmployee(BaseForm):
    """Форма добавления сотрудника"""
    def put_widgets(self):
        # Ввод фамилии
        self.last_name = ttk.Label(self, text="Фамилия*")
        self.last_name.grid(column=0, row=0, sticky='w', cnf=self.master.conf)

        self.entry_last_name = ttk.Entry(self)
        self.entry_last_name.grid(column=1, row=0,
                                  sticky='e', cnf=self.master.conf)

        # Ввод имени
        self.first_name = ttk.Label(self, text="Имя*")
        self.first_name.grid(column=0, row=1, sticky='w', cnf=self.master.conf)

        self.entry_first_name = ttk.Entry(self)
        self.entry_first_name.grid(column=1, row=1,
                                   sticky='e', cnf=self.master.conf)

        # Ввод отчества
        self.patronymic = ttk.Label(self, text="Отчество*")
        self.patronymic.grid(column=0, row=2, sticky='w', cnf=self.master.conf)

        self.entry_patronymic = ttk.Entry(self)
        self.entry_patronymic.grid(column=1, row=2,
                                   sticky='e', cnf=self.master.conf)

        # Ввод табельного номера
        self.time_sheet_number = ttk.Label(self, text="Табельный номер*")
        self.time_sheet_number.grid(column=0, row=3,
                                    sticky='w', cnf=self.master.conf)

        self.entry_time_sheet_number = ttk.Entry(self)
        self.entry_time_sheet_number.grid(
            column=1, row=3,
            sticky='e', cnf=self.master.conf
        )

        # Выбор должности
        self.position = ttk.Label(self, text="Выберите должность*")
        self.position.grid(column=0, row=4, sticky='w', cnf=self.master.conf)

        self.combo_position = ttk.Combobox(self, values=POSITIONS)
        self.combo_position.grid(
            column=1, row=4, sticky='e', cnf=self.master.conf
        )

        # Выбор совместительства
        self.internal_combine = ttk.Label(
            self, text="Сотрудник является совместителем*"
        )
        self.internal_combine.grid(
            column=0, row=5, sticky='w', cnf=self.master.conf
        )

        self.combo_internal_combine = ttk.Combobox(self, values=COMBINE)
        self.combo_internal_combine.grid(
            column=1, row=5, sticky='w', cnf=self.master.conf
        )

        # Выбор ставки
        self.part_time = ttk.Label(self, text="Выберите ставку*")
        self.part_time.grid(column=0, row=6, sticky='w', cnf=self.master.conf)

        self.combo_part_time = ttk.Combobox(self, values=PART_TIME)
        self.combo_part_time.grid(
            column=1, row=6, sticky='e', cnf=self.master.conf
        )

        # Выбор отдела
        self.department = ttk.Label(self, text="Выберите отдел*")
        self.department.grid(column=0, row=7, sticky='w', cnf=self.master.conf)

        self.combo_department = ttk.Combobox(self, values=DEPARTMENTS)
        self.combo_department.grid(
            column=1, row=7, sticky='e', cnf=self.master.conf
        )

        # Кнопка добавления сотрудника
        self.btn = ttk.Button(
            self, text="Добавить", command=self.add_employee
        )
        self.btn.grid(
            column=0, row=8, columnspan=2, sticky='n', padx=10, pady=10
        )

    def add_employee(self):
        # Добавляет сотрудника в бд
        last_name = self.entry_last_name.get()
        first_name = self.entry_first_name.get()
        patronymic = self.entry_patronymic.get()
        time_sheet_number = self.entry_time_sheet_number.get()
        position = self.combo_position.get()
        part_time = self.combo_part_time.get()
        if self.combo_internal_combine.get == '-нет-':
            internal_combine = '*'
        else:
            internal_combine = self.combo_internal_combine.get()
        department = self.combo_department.get()
        employment_date = None
        dismissal_date = None
        employee = search_employee_in_db(last_name, first_name, patronymic,
                                         time_sheet_number, position,
                                         part_time, internal_combine,
                                         department)
        if not check_employee(employee):
            add_employee_to_db(last_name, first_name, patronymic,
                               time_sheet_number, position, part_time,
                               internal_combine, employment_date,
                               dismissal_date, department)
            messagebox.showinfo("Сообщение",  "Сотрудник успешно добавлен!")
        else:
            messagebox.showinfo("Сообщение",  "Сотрудник уже есть в базе!")


class FormDelEmployee(BaseForm):
    """Форма удаления сотрудника"""
    def put_widgets(self):
        # Ввод фамилии удаляемого сотрудника
        self.last_name_employee = ttk.Label(
            self, text="Введите фамилию сотрудника"
        )
        self.last_name_employee.grid(
            column=0, row=0, sticky='w', cnf=self.master.conf
        )

        self.entry_last_name_employee = ttk.Entry(self)
        self.entry_last_name_employee.grid(column=1, row=0,
                                           sticky='e', cnf=self.master.conf)

        # Кнопка, которая запускает поиск сотрудника в базе
        self.btn = ttk.Button(
            self, text="Найти", command=self.search_employee
        )
        self.btn.grid(column=3, row=0, sticky='w', cnf=self.master.conf)

    def del_employee(self):
        employee = self.combo_choice_employee.get().split(', ')
        full_name = employee[0].split()

        del_employee_from_db(full_name[0], full_name[1], full_name[2],
                             employee[1], employee[3].split()[0])
        self.destroy()
        messagebox.showinfo("Сообщение",  "Сотрудник успешно удален из базы!")

    def search_employee(self):
        """Ищет сотрудников с необходимой фамилией"""
        last_name = self.entry_last_name_employee.get()
        self.employees = search_last_name(last_name)
        if self.employees:
            self.text = ttk.Label(
                self, text='Какого сотрудника вы хотите удалить?'
            )
            self.text.grid(column=0, row=1, sticky='n')
            self.EMPLOYEES = [
                f"{empl[0]} {empl[1]} {empl[2]}, {empl[3]}, {empl[4]}, {empl[5]} ст." for empl in self.employees
            ]
            # Выбор сотрудника из найденных
            self.choice_employee = ttk.Label(self, text="Выберите сотрудника")
            self.choice_employee.grid(
                column=0, row=2, sticky='w', cnf=self.master.conf
            )

            self.combo_choice_employee = ttk.Combobox(
                self, values=self.EMPLOYEES, width=60
            )
            self.combo_choice_employee.grid(
                column=1, row=2, sticky='e', cnf=self.master.conf
            )
            self.btn_del = ttk.Button(
                self, text="Удалить", command=self.del_employee
            )
            self.btn_del.grid(
                column=0, row=10, columnspan=2, sticky='n', padx=10, pady=10
            )
        else:
            messagebox.showinfo(
                "Сообщение", "Сотрудник не найден в базе данных!"
            )


class FormTimeSheet(BaseForm):
    """Форма заполнения табеля"""
    def put_widgets(self):
        self.months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
        self.full_month = BooleanVar(0)
        self.file_name_absence = None

        # Ввод года в фомате ХХХХ
        self.year = ttk.Label(self, text="Год в формате XXXX*")
        self.year.grid(column=0, row=0, sticky='w', cnf=self.master.conf)

        self.entry_year = ttk.Entry(self)
        self.entry_year.grid(column=1, row=0,
                             sticky='e', cnf=self.master.conf)

        # Выбор месяца
        self.month = ttk.Label(self, text="Месяц*")
        self.month.grid(column=0, row=1, sticky='w', cnf=self.master.conf)

        self.combo_month = ttk.Combobox(self, values=self.months)
        self.combo_month.grid(column=1, row=1,
                              sticky='e', cnf=self.master.conf)

        # Выбор файла отсутсвия сотрудников
        self.file_absence = ttk.Label(
            self,
            text="Выберите файл отсутствия сотрудников"
        )
        self.file_absence.grid(column=0, row=2,
                               sticky='w', cnf=self.master.conf)

        self.btn_file_absence = ttk.Button(self, text="Выбрать файл",
                                           command=self.choose_file)
        self.btn_file_absence.grid(column=1, row=2,
                                   sticky='e', cnf=self.master.conf)

        # Выбор даты заполнения табеля
        self.filling_date = ttk.Label(self, text="Дата заполнения табеля")
        self.filling_date.grid(column=0, row=4,
                               sticky='w', cnf=self.master.conf)

        self.dateentry_filling_date = DateEntry(
            self, date_pattern='dd.mm.YYYY'
        )
        self.dateentry_filling_date.grid(column=1, row=4,
                                         sticky='e', cnf=self.master.conf)

        # Выбор заполнения: за полный месяц или половину
        self.full_or_no_month = ttk.Label(
            self, text="Заполнять табель за весь месяц?"
        )
        self.full_or_no_month.grid(column=0, row=5,
                                   sticky='w', cnf=self.master.conf)

        self.employment_date_or_not = Radiobutton(
            self, text="да", variable=self.full_month, value=1
        )
        self.employment_date_or_not.grid(column=1, row=5,
                                         sticky='e', cnf=self.master.conf)
        
        # Выбор директории для сохранения
        self.file_absence = ttk.Label(
            self,
            text="Выберите директорию для сохранения табеля"
        )
        self.file_absence.grid(column=0, row=6,
                               sticky='w', cnf=self.master.conf)
        
        self.btn_choose_dir = ttk.Button(
            self,
            text="Выбрать директорию",
            command=self.choose_directory
        )
        self.btn_choose_dir.grid(column=1, row=6, sticky='e', cnf=self.master.conf)

        # Кнопка после ввода всех данных
        self.btn = ttk.Button(
            self, text="Заполнить", command=self.filling_time_sheet
        )
        self.btn.grid(
            column=0, row=10, columnspan=2, sticky='n', padx=10, pady=10
        )

    def choose_file(self):
        """Исполняется при нажатии кнопки
         выбора файла отутствия сотрудниколв"""
        filetypes = (("Таблица", "*.xls, *.xlsx"),)
        filename = filedialog.askopenfilename(
            title="Открыть файл", initialdir="/",
            filetypes=filetypes
        )
        if filename:
            self.file_name = ttk.Label(
                self,
                text=filename
            )
            self.file_name.grid(column=1, row=3,
                                sticky='e', cnf=self.master.conf)
        self.file_name_absence = filename


    def choose_directory(self):
        """Позволяет пользователю выбрать директорию для сохранения файла"""
        directory = filedialog.askdirectory()
        if directory:
            self.save_directory = directory
            ttk.Label(self, text=directory).grid(column=1, row=7,
                                                 sticky='e', cnf=self.master.conf)

    def filling_time_sheet(self):
        """Обрабатывает введенные пользователем данные
        и создает файл с табелем"""
        year = int(self.entry_year.get())
        month = int(self.combo_month.get())
        file_name_absence = self.file_name_absence
        filling_date = self.dateentry_filling_date.get()
        if self.full_month.get() == 1:
            full_month = True
        else:
            full_month = False
        if self.save_directory:
            output_path = self.save_directory
            main_program(year, month, file_name_absence, full_month, filling_date, output_path)
            messagebox.showinfo("Сообщение",  f"Табель успешно создан и сохранен в директории {output_path}!")
        else:
            messagebox.showerror("Ошибка", "Не выбрана директория для сохранения файла!")

if __name__ == "__main__":
    app = App()
    app.mainloop()
