from loguru import logger
import locale
import os
from conf import *
import PySimpleGUI as sg

from windows.kids_window import kids_window
from windows.add_new_sheet import add_new_sheet
from windows.mark_kids import mark_kids
from windows.notes import notes_window
from windows.settings import settings_window
from utils import close_all_sheets

locale.setlocale(locale.LC_ALL, 'ru-RU')

sg.theme(THEME)


def check_directory():
    if not os.path.exists('Ведомости'):
        os.mkdir('Ведомости')


def main_window(font_family=FONT_FAMILY, font_size=FONT_SIZE):
    # Основное окно приложения, само по себе ничего не делает, по сути навигационное меню.
    # Единственное значение которое можно передать дальше название группы

    # Собираем файлы ведомостей, для добавления в выпадающий список, если файлов нет, создается пустой список
    try:
        list_group = [group for group in os.listdir(path=r'Ведомости/') if '.xlsx' in group]
        default_val = list_group[0]
    except IndexError:
        list_group = []
        default_val = ''

    layout = [
        [sg.Menu([['Настройки', ['Параметры']]], font=(FONT_FAMILY, 12))],
        [sg.Text("Ведомости детский сад")],
        [sg.Button('Добавить ведомость', expand_x=True), sg.Button('Ученики', expand_x=True)],
        [sg.Combo(list_group, expand_x=True, default_value=default_val, key='file_name', readonly=True),
         sg.Button('Отметить')],
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
            kids_window(values['file_name'])
            window.reappear()

        if event == 'Добавить ведомость':
            window.disappear()
            add_new_sheet(values['file_name'], list_group)
            list_group = [group for group in os.listdir(path=r'Ведомости/') if '.xlsx' in group]
            window['file_name'].update(values=list_group, set_to_index=0)
            window.reappear()

        if event == 'Отметить':
            window.disappear()
            mark_kids(values['file_name'])
            window.reappear()

        if event == 'Заметки':
            window.disappear()
            notes_window(values['file_name'])
            window.reappear()

        if event == 'Закрыть ведомости':
            if sg.Window('Вы уверены?', [
                [sg.Text('Закрывать ведомости нужно только в конце месяца!', font=(FONT_FAMILY, 35)),
                 sg.Button('Да', font=(FONT_FAMILY, 35)), sg.Button('Отмена', font=(FONT_FAMILY, 35))]
            ]).read(close=True)[0] == 'Да':
                close_all_sheets()

        if event == 'Параметры':
            window.disappear()
            settings_window()
            window.reappear()

    window.close()


logger.add('debug.log', level='ERROR')
logger.add('absents.log', level='INFO')


@logger.catch()
def main():
    check_directory()
    main_window()


if __name__ == '__main__':
    main()