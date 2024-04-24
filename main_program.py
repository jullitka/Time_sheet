from openpyxl import Workbook
from openpyxl.worksheet.pagebreak import Break, RowBreak

from data_from_user import (department_name, keeps_time_sheet,
                            position, head_of_department,
                            head_of_HR)

from table_drawing import (table_num_date, table_header,
                           form_info, table_report_period,
                           name_company, document_name,
                           table_to_one_emploee, total_sums,
                           signature, signature_date,
                           fill_dates, add_data_to_time_sheet,
                           put_down_dates, row_dimension)
from utils import choose_employee


def main_program(year, month, file_name_absence, full_month, report_date, output_path):

    time_sheet_name = f'{output_path}/ТАБЕЛЬ {month}.{year}.xlsx'
    wb = Workbook()
    ws = wb.active

    # ориентация листа
    ws.page_setup.orientation = 'landscape'
    # размер листа
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    row_break = RowBreak()
    row_break.append(Break(id=5))
    row_break.append(Break(id=90))
    ws.row_breaks = row_break

    # строка, в которой начинается таблица для сотрудника
    num = 18

    # отчет сотрудников начинается с
    count = 1

    # все сотрудники в виде списка
    list_employees = choose_employee()
    # количество сотрудников в табеле
    len_table = len(list_employees)

    table_num_date(time_sheet_name)
    table_header(time_sheet_name)
    form_info(time_sheet_name)
    table_report_period(time_sheet_name)
    name_company(department_name, time_sheet_name)
    document_name(time_sheet_name)
    for i in range(len_table):
        employee = list_employees[i]
        add_data_to_time_sheet(
            employee, num, file_name_absence,
            year, month, full_month, time_sheet_name
        )
        table_to_one_emploee(num, count, time_sheet_name)
        num += 4
        count += 1
    total_sums(num, time_sheet_name)
    signature(
        num, position, keeps_time_sheet,
        head_of_department, head_of_HR, time_sheet_name
    )
    signature_date(num, time_sheet_name)
    fill_dates(year, month, time_sheet_name)
    put_down_dates(year, month, report_date, time_sheet_name)
    row_dimension(num, time_sheet_name)
