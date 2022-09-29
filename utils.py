from openpyxl import load_workbook
from pathlib import Path
from datetime import datetime
from calendar import Calendar
from conf import *

MONTH_DICT = {
    'Сентябрь': (9, YEAR[0]), 'Октябрь': (10, YEAR[0]), 'Ноябрь': (11, YEAR[0]), 'Декабрь': (12, YEAR[0]),
    'Январь': (1, YEAR[1]), 'Февраль': (2, YEAR[1]), 'Март': (3, YEAR[1]), 'Апрель': (4, YEAR[1]), 'Май': (5, YEAR[1])
}

date = datetime.now()


def get_work_days(month):
    wd = (2, 4)  # Рабочие дни (СРЕДА, ПЯТНИЦА)
    return [day[0] for day in Calendar().itermonthdays2(MONTH_DICT[month][1], MONTH_DICT[month][0])
            if day[0] != 0 and day[1] in wd]


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


if __name__ == '__main__':
    pass
