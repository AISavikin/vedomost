from gui import *
from loguru import logger
import locale
import os
from conf import *

locale.setlocale(locale.LC_ALL, 'ru-RU')

sg.theme(THEME)

def check_directory():
    if not os.path.exists('Ведомости'):
        os.mkdir('Ведомости')


def main_window(font_family=FONT_FAMILY, font_size=FONT_SIZE, preview=False):
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
        [sg.Button('Добавить ведомость', expand_x=True), sg.Button('Добавить ученика')],
        [sg.Combo(list_group, expand_x=True, default_value=default_val, key='file_name', readonly=True),
         sg.Button('Отметить')],
        [sg.Button('Заметки', expand_x=True)]
    ]

    window = sg.Window('Ведомости', layout, element_justification='center', font=(font_family, font_size))

    if preview:
        while True:
            event, values = window.read()
            if event == sg.WINDOW_CLOSED:
                break
    else:
        while True:
            event, values = window.read()
            if event == sg.WINDOW_CLOSED:
                break

            if event == 'Добавить ученика':
                add_new_kids(values['file_name'])

            if event == 'Добавить ведомость':
                add_new_sheet(values['file_name'], list_group)
                list_group = [group for group in os.listdir(path=r'Ведомости/') if '.xlsx' in group]
                window['file_name'].update(values=list_group, set_to_index=0)

            if event == 'Отметить':
                check_kids(values['file_name'])

            if event == 'Заметки':
                notes_window(values['file_name'])

            if event == 'Параметры':
                window.disappear()
                settings_window()
                window.reappear()

    window.close()


logger.add('debug.log')


@logger.catch()
def main():
    check_directory()
    main_window()


if __name__ == '__main__':
    main()
