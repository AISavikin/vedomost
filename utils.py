from datetime import datetime
from calendar import Calendar
from conf import *
from database import Student

MONTH_DICT = {
    'Сентябрь': (9, YEAR[0]), 'Октябрь': (10, YEAR[0]), 'Ноябрь': (11, YEAR[0]), 'Декабрь': (12, YEAR[0]),
    'Январь': (1, YEAR[1]), 'Февраль': (2, YEAR[1]), 'Март': (3, YEAR[1]), 'Апрель': (4, YEAR[1]), 'Май': (5, YEAR[1])
}

date = datetime.now()

def get_month_name(num_month):
    for key, val in MONTH_DICT.items():
        if val[0] == num_month:
            return key

def get_work_days(month):
    wd = (2, 4)  # Рабочие дни (СРЕДА, ПЯТНИЦА)
    return [day[0] for day in Calendar().itermonthdays2(MONTH_DICT[month][1], MONTH_DICT[month][0])
            if day[0] != 0 and day[1] in wd]


def get_kids(num_group):
    """
    :param num_group: Номер группы из базы данных.
    :return: Список строк с именами детей
    """
    return [student for student in Student.select().where(Student.group == num_group).order_by(Student.name)]



