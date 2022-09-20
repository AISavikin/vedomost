import os
from openpyxl import load_workbook
import shutil


def check_directory():
    if not os.path.exists('Ведомости'):
        os.mkdir('Ведомости')


def write_xls(func):
    """
    Декоратор, для открытия и записи файла Excel.
    Первым параметром, обязательно передавать относительный путь к файлу
    """

    def wrapper(file_name: str, *args):
        wb = load_workbook(f'Ведомости/{file_name}')
        ws = wb.active

        func(ws, *args)

        wb.save(f'Ведомости/{file_name}')

    return wrapper


def get_kids(file_name):
    """
    :param file_name: Относительный путь к файлу
    :return: Список строк с именами детей
    """
    wb = load_workbook(f'Ведомости/{file_name}')
    ws = wb.active
    names = ws['B16:B38']
    return [names[i][0].value for i in range(23) if names[i][0].value]


@write_xls
def save_new_kids(ws, kids):
    cells = 'B16:B38'
    for i in range(len(kids)):
        ws[cells][i][0].value = kids[i]


@write_xls
def clear_absent(ws):
    # Очищает ячейки от "н, б"
    # 16 - 39 диапазон строк, 5 - 36 диапазон столбцов
    for row in range(16, 39):
        for col in range(5, 36):
            ws.cell(row=row, column=col).value = ''


@write_xls
def check_absent(ws, day, absent):
    column = day + 4
    for row in range(len(absent)):
        ws.cell(row=row + 16, column=column).value = absent[row]

@write_xls
def write_service_information(ws, month, group):
    ws['N3'].value = month
    ws['AA42'].value = month
    ws['C5'].value = group


def copy_sheet(file_name, base, month):
    if base == 'Новая группа':
        shutil.copyfile('Template.xlsx', f'Ведомости/{file_name}_{month}.xlsx')
        return f'{file_name}_{month}.xlsx'
    file_name = base.split('_')[0]
    new_path = f'Ведомости/{file_name}_{month}.xlsx'
    shutil.copyfile(f'Ведомости/{base}', new_path)
    return f'{file_name}_{month}.xlsx'

def create_new_sheet(file_name, base, month):

    file_name = copy_sheet(file_name, base, month)
    group = file_name.split('_')[0]
    clear_absent(file_name)
    write_service_information(file_name, month, group)


def test():
    file_name = 'Группа 1'
    base = 'Новая группа'
    month = 'Октябрь'
    create_new_sheet(file_name, base, month)

if __name__ == '__main__':
    test()
