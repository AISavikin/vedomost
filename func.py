import calendar
import os
from openpyxl import load_workbook, styles
import shutil

MONTH_DICT = {
    'Сентябрь': 9, 'Октябрь': 10, 'Ноябрь': 11, 'Декабрь': 12, 'Январь': 1, 'Февраль': 2, 'Март': 3, 'Апрель': 4,
    'Май': 5
}

def check_directory():
    if not os.path.exists('Ведомости'):
        os.mkdir('Ведомости')




def get_kids(file_name):
    """
    :param file_name: Относительный путь к файлу
    :return: Список строк с именами детей
    """
    wb = load_workbook(f'Ведомости/{file_name}')
    ws = wb['Посещаемость']
    names = ws['B16:B38']
    return [names[i][0].value for i in range(23) if names[i][0].value]


def save_new_kids(file_name, kids):
    file_name = f'Ведомости/{file_name}'
    work_book = load_workbook(file_name)
    ws = work_book['Посещаемость']
    cells = 'B16:B38'

    for i in ws[cells]:
        i[0].value = ''

    for i in range(len(kids)):
        ws[cells][i][0].value = kids[i]

    work_book.save(file_name)


def clear_absent(file_name):
    file_name = f'Ведомости/{file_name}'
    # Очищает ячейки от "н, б"
    # 16 - 39 диапазон строк, 5 - 36 диапазон столбцов
    work_book = load_workbook(file_name)
    ws = work_book['Посещаемость']
    for row in range(16, 39):
        for col in range(5, 36):
            ws.cell(row=row, column=col).value = ''
            ws.cell(row=row, column=col).fill = styles.PatternFill()

    work_book.save(file_name)


def mark_absent(file_name, day, absent):
    file_name = f'Ведомости/{file_name}'
    work_book = load_workbook(file_name)
    ws = work_book['Посещаемость']
    column = day + 4
    for row in range(len(absent)):
        ws.cell(row=row + 16, column=column).value = absent[row]

    work_book.save(file_name)


def write_service_information(file_name, month, group):
    file_name = f'Ведомости/{file_name}'
    work_book = load_workbook(file_name)
    ws = work_book['Посещаемость']
    ws['N3'].value = month
    ws['AA42'].value = month
    ws['C5'].value = group

    work_book.save(file_name)


def colorize_weekend(file_name, month):
    file_name = f'Ведомости/{file_name}'
    work_book = load_workbook(file_name)
    ws = work_book['Посещаемость']

    c = calendar.Calendar()
    weekends = [day[0] for day in c.itermonthdays2(2022, MONTH_DICT[month]) if day[0] != 0 and day[1] in (5, 6)]
    for row in range(16, 39):
        for col in weekends:
            ws.cell(row=row, column=col + 4).fill = styles.PatternFill(start_color='5E5E5E', fill_type='solid')

    work_book.save(file_name)


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
    colorize_weekend(file_name, month)


def save_notes(file_name, day, note):
    """!!!!!!!!!!!!ДОПИЛИТЬ!!!!!!!!!!!"""

    wb = load_workbook(f'Ведомости/{file_name}')
    ws = wb['Заметки']
    row = 1
    while ws.cell(row=row, column=1).value:
        if ws.cell(row=row, column=1).value == day:
            break
        row += 1
    ws.cell(row=row, column=1).value = day
    ws.cell(row=row, column=2).value = note
    wb.save(f'Ведомости/{file_name}')


def read_note(file_name, day):
    """!!!!!!!!!!!!ДОПИЛИТЬ!!!!!!!!!!!"""

    wb = load_workbook(f'Ведомости/{file_name}')
    ws = wb['Заметки']
    row = 1
    while ws.cell(row=row, column=1).value:
        if ws.cell(row=row, column=1).value == day:
            break
        row += 1
    if not ws.cell(row=row, column=2).value:
        return ''
    return ws.cell(row=row, column=2).value


def clear_note(file_name):
    """!!!!!!!!!!!!ДОПИЛИТЬ!!!!!!!!!!!"""
    wb = load_workbook(f'Ведомости/{file_name}')
    ws = wb['Заметки']
    cells = 'A1:A20'
    for i in ws[cells]:
        i[0].value = ''
    cells = 'B1:B20'
    for i in ws[cells]:
        i[0].value = ''

    wb.save(f'Ведомости/{file_name}')


if __name__ == '__main__':
    print(get_kids('Группа 1_Октябрь.xlsx'))
