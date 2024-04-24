from datetime import date, datetime
import calendar
import sqlite3

from information import ADDITIONAL_HOLIDAYS, SHORT_DAYS, WORKING_SATYRDAYS 


def month_name_russian(month:int):
    if month == 1:
        return 'Январь'
    elif month == 2:
        return 'Февраль'
    elif month == 3:
        return 'Март'
    elif month == 4:
        return 'Апрель'
    elif month == 5:
        return 'Май'
    elif month == 6:
        return 'Июнь'
    elif month == 7:
        return 'Июль'
    elif month == 8:
        return 'Август'
    elif month == 9:
        return 'Сентябрь'
    elif month == 10:
        return 'Октябрь'
    elif month == 11:
        return 'Ноябрь'
    elif month == 12:
        return 'Декабрь'



def check_date_exist(year, month, day):
    """Проверяет, существует ли дата"""
    try:
        return date(year, month, day)
    except ValueError:
        return False


def check_holiday(year, month, day):
    """Проверяет, является ли день выходным
    и возвращает True в этом случае"""
    if check_date_exist(year, month, day):
        my_date = date(year, month, day)
    else:
        return False
    if my_date.strftime('%d.%m.%Y') in WORKING_SATYRDAYS:
        return False
    if my_date.weekday() in (5, 6):
        return True
    if my_date.strftime('%d.%m.%Y') in ADDITIONAL_HOLIDAYS:
        return True
    return False


def check_short_day(year, month, day):
    """Проверяет, является ли день коротким
    и возвращает True в этом случае"""
    my_date = date(year, month, day)
    if my_date.strftime('%d.%m.%Y') in SHORT_DAYS:
        return True
    else:
        return False


def duration(year, month, day, part_time):
    """Определяет продолжительность дня для данной ставки и дня"""
    if check_short_day(year, month, day):
        return part_time*8 - 1
    else:
        return part_time*8


def start_or_end_filling(year, month, key_date, dismissal=False):
    """Определяет, входит ли ключевая дата в месяц.
    Актуально в случае недавнего трудоустройства или увольнения
    Возвращает дату"""
    # первый день месяца
    first_date = datetime(year, month, 1)
    # последний день месяца
    last_date = datetime(year, month, calendar.monthrange(year, month)[1])
    # преобразуем строку с датой устройства в дату
    if key_date is not None:
        key_date = datetime.strptime(key_date, '%Y-%m-%d %H:%M:%S')
        if first_date <= key_date <= last_date:
            return key_date
        elif dismissal is True:
            return last_date
        else:
            return first_date
    else:
        if dismissal is True:
            return last_date
        else:
            return first_date


def reporting_period(year, month, report_date):
    """Определяет за какой период заполнять табель
    используя дату заполнения.
    Возвращет дату окончания периода в виде строки"""
    report_date = datetime.strptime(report_date, '%d.%m.%Y')
    if abs(
        datetime(year, month, 15) - report_date
        ) < abs(
            datetime(year, month, calendar.monthrange(year, month)[1]) - report_date
            ):
        end_date = datetime(year, month, 15).strftime('%d.%m.%Y')
        return end_date
    else:
        end_date = datetime(
            year, month, calendar.monthrange(year, month)[1]
        ).strftime('%d.%m.%Y')
        return end_date


def choose_employee():
    """Выбирает из базы всех сотрудников, отсортированных в
    алфавитном порядке и возвращает их в виде списка кортежей"""
    conn = sqlite3.connect('employees.db')
    cur = conn.cursor()
    cur.execute(
        """SELECT * FROM employees
        ORDER BY last_name""")
    list_employees = cur.fetchall()
    conn.commit()
    return list_employees


def check_employee(employee):
    """Проверяет, есть ли сотрудник в базе данных.
    Возвращает True если есть."""
    if employee is None:
        return False
    else:
        return True


def generate_name(last_name, first_name, patronymic):
    """Генерирует ФИО в формате 'Фамилия И.О.'"""
    return last_name + ' ' + first_name[0] + '.' + patronymic[0] + '.'


def letter_code(string):
    """Сопоставляет причине отсутствия буквенный код"""
    if string.lower() in ('болезнь', 'больничный'):
        return 'Б'
    elif string.lower() in ('отпуск основной',):
        return 'ОТ'
    elif string.lower() == 'командировка': #in ('командировка',):
        return 'К'
    elif string.lower() in (
        'отпуск неоплачиваемый по разрешению работодателя',
    ):
        return 'ДО'
    elif string.lower() in ('отпуск по уходу за ребенком',):
        return 'ОЖ'
    elif string.lower() in ('отстранение от работы с оплатой',):
        return 'НО'
    elif string.lower() in ('дополнительный отпуск',):
        return 'ОД'
    elif string.lower() in ('отпуск по беременности и родам',):
        return 'Р'
    elif string.lower() in ('дополнительные выходные дни (неоплачиваемые)',):
        return 'НВ'
    elif string.lower() in ('дополнительные выходные дни (оплачиваемые)',):
        return 'ОВ'
    elif string.lower() in ('отсутствие по невыясненным причинам',):
        return 'НН'
    else:
        return '???'

