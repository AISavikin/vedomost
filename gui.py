import shutil
import PySimpleGUI as sg
from func import *
from datetime import datetime

# add_group_window = sg.Window('Добавить группу',
#                              [[sg.Text('Введите назывние группы'), sg.Input(size=(30, 10)), sg.OK()]],
#                              finalize=True, element_justification='right')


def main_window():
    list_group = [group.split('.')[0] for group in os.listdir() if '.xls' in group]
    layout = [
        [sg.Text("Ведомости детский сад", key='lable')],
        [sg.Button('Добавить группу'), sg.Button('Добавить ученика')],
        [sg.Combo(list_group, expand_x=True, default_value=list_group[0], key='combo'), sg.Button('Отметить')]
    ]

    window = sg.Window('Ведомости', layout, element_justification='center')
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'Quit':
            break
        if event == 'Добавить ученика':
            add_kid(values['combo'])
        if event == 'Добавить группу':
            # add_group_window.read(close=True)
            add_group_window()
        if event == 'Отметить':
            check_kids(values['combo'])

    window.close()


def add_group_window():
    layout = [
        [sg.Text('Введите назывние группы'), sg.Input(size=(30, 10)), sg.OK()],
    ]
    window = sg.Window('Добавить группу', layout, finalize=True, element_justification='right')
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'Quit':
            break
        elif event == 'OK':
            shutil.copyfile('Template.xls', f'Ведомости/{values[0]}.xls')
            break
    window.close()


def add_kid(group):
    layout = [
        [sg.Text('Имя и фамилия ученика'), sg.Input(size=(30, 10)), sg.OK()],
    ]
    window = sg.Window('Добавить ученика', layout, finalize=True, element_justification='right')
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == 'OK':
            row = len(read_xls(group))
            write_xls([values[0]], group, row=row)
    window.close()


def check_kids(file_name):
    date = datetime.now()
    date = int(f'{date: %d}')
    kids = read_xls(file_name)
    layout = [
        [sg.Text('Дата: ', auto_size_text=True), sg.Combo([i for i in range(1, 32)], default_value=date, key='date')],
        [],
        [sg.OK()]
    ]
    layout[1] = [[sg.Text(kid), sg.Combo(('н', 'б'), default_value='б')] for kid in kids]
    window = sg.Window('Отметить ученика', layout, finalize=True, element_justification='right')
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == 'OK':
            day = values['date']
            data = [values[i] for i in range(len(kids))]
            write_xls(data, file_name, day)

    window.close()


def main():
    check_directory()
    main_window()


if __name__ == '__main__':
    main()
