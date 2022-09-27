import PySimpleGUI as sg
from calendar import Calendar
from func import MONTH_DICT
from openpyxl import load_workbook

lorem = """Lorem ipsum dolor  auctor, vestibulum lacinia velit. Sed dapibus vel dolor et semper.
"""


def notes_window(file_name: str):
    notes = read_notes(file_name)

    # notes = {
    #     2: lorem,
    #     7: lorem,
    #     9: lorem,
    #     14: lorem,
    #     16: lorem,
    #     21: lorem,
    #     23: lorem,
    #     28: lorem,
    #     30: lorem,
    #
    # }

    n = [[sg.Text(day),
          sg.Multiline(default_text=notes[day], size=(50, 6), key=day, no_scrollbar=True, auto_size_text=True)] for day
         in notes]

    layout = [
        [sg.Text(file_name)],
        [sg.Column(n, scrollable=True, vertical_scroll_only=True, size_subsample_width=1, element_justification='r',
                   size=(400, 600), key='column')],
        [sg.Button('Сохранить')]
    ]

    window = sg.Window('Заметки', layout, element_justification='center', modal=True)

    while True:
        event, values = window.read()

        if event == 'Отмена' or event == sg.WINDOW_CLOSED:
            break

        if event == 'Сохранить':
            notes = {day: values[day] for day in values if day != 'file_name'}
            save_notes(file_name, notes)
            break


def get_work_days(month, year=2022, wd=(2, 4)):
    c = Calendar()
    return [day[0] for day in c.itermonthdays2(year, month) if day[0] != 0 and day[1] in wd]


def read_notes(file_name):
    workbook = load_workbook(f'Ведомости/{file_name}')
    ws = workbook['Заметки']
    row = 1
    notes = {}
    while ws.cell(row=row, column=1).value:
        notes[ws.cell(row=row, column=1).value] = ws.cell(row=row, column=2).value
        row += 1
    return notes


def save_notes(file_name, notes):
    file_name = f'Ведомости/{file_name}'
    workbook = load_workbook(file_name)
    ws = workbook['Заметки']
    row = 1
    for day in notes:
        ws.cell(row=row, column=1).value = day
        ws.cell(row=row, column=2).value = notes[day]

    workbook.save(file_name)

def write_service_information(file_name, month, group):
    file_name = f'Ведомости/{file_name}'
    work_book = load_workbook(file_name)
    ws = work_book['Посещаемость']
    ws['N3'].value = month
    ws['AA42'].value = month
    ws['C5'].value = group
    ws = work_book['Заметки']
    work_days = get_work_days(MONTH_DICT[month])
    row = 1
    for day in work_days:
        ws.cell(row=row, column=1).value = day
        row += 1


    work_book.save(file_name)


def main():
    print(get_work_days(9))
    print(read_notes('Группа 1_Ноябрь.xlsx'))
    write_service_information('Группа 1_Ноябрь.xlsx', 'Ноябрь', 'НАЗВАНИЕ ГРУППЫ')

if __name__ == '__main__':
    main()
