import os
from openpyxl import load_workbook
import shutil


def write_xls(func):
    def wrapper(file_name, *args):
        wb = load_workbook(f'Ведомости/{file_name}')
        ws = wb.active

        res = func(ws, *args)

        wb.save(f'Ведомости/{file_name}')

    return wrapper


def get_kids(file_name):
    wb = load_workbook(f'Ведомости/{file_name}')
    ws = wb.active
    names = ws['B16:B38']
    return [names[i][0].value for i in range(23) if names[i][0].value]


@write_xls
def save_new_kids(ws, *args):
    cells = 'B16:B38'
    kids = args[0]
    for i in range(len(kids)):
        ws[cells][i][0].value = kids[i]


@write_xls
def clear_absent(ws):
    for row in range(16, 39):
        for col in range(5, 36):
            ws.cell(row=row, column=col).value = '*'


def create_new_sheet(file_name, month, base):
    print(f'{file_name=}\n{base=}\n{month=}')
    if base == 'Новая группа':
        shutil.copyfile('Template.xlsx', f'Ведомости/{file_name}_{month}.xlsx')
        return
    file_name = base.split('_')[0]
    shutil.copyfile(f'Ведомости/{base}', f'Ведомости/{file_name}_{month}.xlsx')


def check_absent(file_name, absent):
    pass


def check_directory():
    if not os.path.exists('Ведомости'):
        os.mkdir('Ведомости')


def test():
    pass


if __name__ == '__main__':
    clear_absent('Группа 1_Апрель.xlsx')
