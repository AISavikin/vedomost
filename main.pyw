from loguru import logger
import locale
import os
from conf import *
import PySimpleGUI as sg
from database import create_database, Student

from windows.kids_window import kids_window
from windows.sheet_window import sheet
from windows.add_new_sheet import add_new_sheet
from windows.mark_kids import mark_kids
from windows.notes import notes_window
from windows.settings import settings_window
from utils import close_all_sheets, MONTH_DICT, date, get_month_name

# locale.setlocale(locale.LC_ALL, 'ru-RU')

sg.theme(THEME)


def check_directory():
    """Проверяет существует ли директория, и создает если необходимо"""
    if not os.path.exists('Ведомости'):
        os.mkdir('Ведомости')


def main_window(font_family=FONT_FAMILY, font_size=FONT_SIZE):
    # Основное окно приложения, само по себе ничего не делает, по сути навигационное меню.
    # Единственное значение которое можно передать дальше номер группы

    # Собираем номера групп из базы данных, если база пустая, то создается пустой список.
    list_group = [f'Группа {num}' for num in set(i.group for i in Student.select())]
    try:
        default_val = list_group[0]
    except IndexError:
        default_val = ''

    layout = [
        [sg.Menu([['Настройки', ['Параметры']]], font=(FONT_FAMILY, 12))],
        [sg.Text("Ведомости детский сад")],
        [sg.Button('Добавить группу', expand_x=True), sg.Button('Ученики', expand_x=True)],
        [sg.Combo(list_group, expand_x=True, default_value=default_val, key='num_group', readonly=True),
         sg.Combo([month_name for month_name in MONTH_DICT], default_value=get_month_name(date.month), key='month_name')],
        [sg.Button('Отметить', expand_x=True)],
        [sg.Button('Закрыть ведомости', expand_x=True)],
        [sg.Button('Заметки', expand_x=True)]
    ]

    window = sg.Window('Ведомости', layout, element_justification='center', font=(font_family, font_size))

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break

        if event == 'Ученики':
            window.disappear()
            kids_window(values['num_group'].split()[-1])
            window.reappear()

        if event == 'Добавить группу':
            window.disappear()
            num_group = sg.PopupGetText('Введите номер группы', size=(10, 10))
            if not num_group:
                window.reappear()
                continue
            if not num_group.isdigit():
                sg.Popup('Только цифры!')
                window.reappear()
                continue
            kids_window(int(num_group))
            window.reappear()

        if event == 'Отметить':
            window.disappear()
            mark_kids(values['num_group'].split()[-1], values['month_name'])
            window.reappear()

        # if event == 'Заметки':
        #     window.disappear()
        #     notes_window(values['file_name'])
        #     window.reappear()

        if event == 'Закрыть ведомости':
            window.disappear()
            sheet(int(values['num_group'].split()[-1]), values['month_name'])
            window.reappear()

        if event == 'Параметры':
            window.disappear()
            settings_window()
            window.reappear()

    window.close()


logger.add('logs/debug.log', level='ERROR')
logger.add('logs/absents.log', level='INFO')


def main():
    check_directory()
    create_database()
    main_window()


if __name__ == '__main__':
    main()
