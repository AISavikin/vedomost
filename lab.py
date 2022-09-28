import PySimpleGUI as sg
from conf import *
from main import main_window

lorem = """Lorem ipsum dolor  auctor, vestibulum lacinia velit. Sed dapibus vel dolor et semper.
"""


def settings_window():
    year, font_family, font_size, theme = YEAR, FONT_FAMILY, FONT_SIZE, THEME

    layout = [
        [sg.Text('Год:'), sg.Combo(list(range(2022, 2031)), default_value=year, enable_events=True, k='year')],
        [sg.Text('Тема: '), sg.Combo(sg.theme_list(), default_value=theme, k='theme')],
        [sg.Button('Предпросмотр темы', expand_x=True)],
        [sg.Button('Изменить шрифт', expand_x=True)],
        [sg.Button('Сохранить'), sg.Button('Отмена', expand_x=True)]
    ]
    window = sg.Window('Заметки', layout, element_justification='center', font=(font_family, font_size))

    while True:
        event, values = window.read()

        if event == 'Отмена' or event == sg.WINDOW_CLOSED:
            break

        if event == 'Предпросмотр темы':
            sg.theme(values['theme'])
            main_window(font_family=font_family, font_size=font_size, preview=True)

        if event == 'year':
            year = values['year']

        if event == 'Изменить шрифт':
            window.disappear()
            font_family, font_size = change_font()
            window.reappear()

        if event == 'Сохранить':
            theme = values['theme']
            with open('conf.py', 'w') as f:
                f.write(f'YEAR = {year}\nFONT_FAMILY = "{font_family}"\nFONT_SIZE = {font_size}\nTHEME = "{theme}"')

    window.close()


def change_font():
    font_family = FONT_FAMILY
    font_size = FONT_SIZE
    fonts = ['Verdana', 'Times New Roman', 'Tahoma', 'Book Antiqua', 'Courier', 'Calibri']
    layout = [
        [sg.Text('Шрифт'), sg.Combo(fonts, default_value=FONT_FAMILY, enable_events=True, k='font'), sg.Button('OK')],
        [sg.Column([[sg.Text('Текст для проверки размера\nи шрифта', k='text')]], size=(950, 180))],
        # [sg.Text('Текст для проверки размера\nи шрифта', k='text')],
        [sg.Slider((6, 50), default_value=font_size, expand_x=True, orientation='h', key='-slider-',
                   enable_events=True)],

    ]

    window = sg.Window('Смена шрифта', layout, font=(FONT_FAMILY, FONT_SIZE), size=(900, 300))
    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED:
            break

        if event == '-slider-':
            font_size = int(values['-slider-'])
            window['text'].update(font=(font_family, font_size))

        if event == 'font':
            font_family = values['font']
            window['text'].update(font=(font_family, font_size))

        if event == 'OK':
            break

    window.close()
    return font_family, font_size


def save_settings(year, font_family, font_size, theme):
    with open('conf.py', 'w') as f:
        f.write(f'YEAR = {year}\nFONT_FAMILY = {font_family}\nFONT_SIZE = {font_size}\nTHEME = {theme}')


def main():
    fonts = ['Verdana', 'Times New Roman', 'Tahoma', 'Book Antiqua', 'Courier', 'Calibri']
    settings_window()


if __name__ == '__main__':
    main()
