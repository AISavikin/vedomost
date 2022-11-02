from conf import *
import PySimpleGUI as sg
from database import Student
from datetime import datetime
from utils import get_kids

def gen_table(num_group):
    cnt = 1
    table = []
    for i in get_kids(num_group):
        table.append([cnt, i.name, f'{i.added:%d.%m.%y}'])
        cnt += 1
    return table

def kids_window(num_group: int):
    """
    Создает окно для добавления новых детей в ведомость.

    :param num_group: int: Номер группы из базы данных
    """

    # Левая колонка
    left_col = [
        [sg.Text('Фамилия и имя')],
        [sg.Input(key='name', focus=True, size=29)],
        [sg.Button('Добавить'), sg.Button('Удалить'), sg.Button('Исправить')]
    ]
    # Правая колонка
    right_col = [
        [sg.Text(f'Группа {num_group}')],
        [sg.Table(headings=['№', 'Имя', 'Добавлен'], values=gen_table(num_group),
                  key='table', col_widths=[7, 20, 10],
                  auto_size_columns=False, justification='center', size=(20, 15), enable_events=True)]
    ]

    layout = [[sg.Column(left_col, element_justification='c'), sg.Column(right_col, element_justification='c')]]
    window = sg.Window(f'Добавить ученика в группу {num_group}', layout, finalize=True,
                       font=(FONT_FAMILY, FONT_SIZE))
    window.bind("<Return>", "Добавить")

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break

        if event == 'table':
            kids = [kid.name for kid in get_kids(num_group)]
            if values['table']:
                window['name'].update(kids[values['table'][0]])

        if event == 'Добавить' and values['name']:
            Student.create(name=values['name'], group=num_group, added=datetime.now(), active=True)
            window['name'].update('')
            window['table'].update(values=gen_table(num_group))

        if event == 'Исправить' and values['table']:
            if not values['name']:
                sg.Popup('Введите имя!')
                continue
            old_name = get_kids(num_group)[values['table'][0]].name
            new_name = values['name']
            Student.update({Student.name: new_name}).where(Student.name == old_name).execute()
            window['table'].update(values=gen_table(num_group))
            window['name'].update('')

        if event == 'Удалить':
            Student.delete().where(Student.name == values['name']).execute()
            window['table'].update(values=gen_table(num_group))
            window['name'].update('')

    window.close()



if __name__ == '__main__':
    kids_window(7)

