from openpyxl import load_workbook
from pathlib import Path
from datetime import datetime
from calendar import Calendar
from conf import *


MONTH_DICT = {
    'Сентябрь': 9, 'Октябрь': 10, 'Ноябрь': 11, 'Декабрь': 12, 'Январь': 1, 'Февраль': 2, 'Март': 3, 'Апрель': 4,
    'Май': 5
}

date = datetime.now()

def get_work_days(month):
    wd = (2, 4)  # Рабочие дни (СРЕДА, ПЯТНИЦА)
    return [day[0] for day in Calendar().itermonthdays2(YEAR, MONTH_DICT[month]) if day[0] != 0 and day[1] in wd]

def get_kids(file_name):
    """
    :param file_name: Относительный путь к файлу
    :return: Список строк с именами детей
    """
    path = Path(Path().cwd(), 'Ведомости', file_name)
    wb = load_workbook(path)
    ws = wb['Посещаемость']
    names = ws['B16:B38']
    return [names[i][0].value for i in range(23) if names[i][0].value]
