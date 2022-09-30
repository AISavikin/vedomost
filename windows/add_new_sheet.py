import PySimpleGUI as sg
from utils import MONTH_DICT, date, create_new_sheet
from conf import *


def add_new_sheet(file_name: str, list_group: list):
    """
    Создает окно для добавления новой ведомости
    :param file_name: str имя файла без расширения
    :param list_group: list список существующих ведомостей
    """
    # Добавляем в список групп значение "Новая группа"
    if 'Новая группа' not in list_group:
        list_group.append('Новая группа')

    left = [
        [sg.Text('На основе')],
        [sg.Text('Название ведомости')],
        [sg.Text('Месяц')]
    ]
    right = [
        [sg.Combo(list_group, size=(18, 0), key='base', default_value=file_name, readonly=True)],
        [sg.Input(size=(18, 0), key='file_name')],
        [sg.Combo(list(MONTH_DICT.keys()), size=(18, 0), key='month', default_value=f'{date:%B}', readonly=True)]
    ]
    layout = [
        [sg.Column(left), sg.Column(right)],
        [sg.Button('Добавить ведомость', key='add'), sg.Button('Отмена')]
    ]

    window = sg.Window('Добавить новую ведомость', layout, element_justification='c', font=(FONT_FAMILY, FONT_SIZE))

    while True:
        event, values = window.read()

        if event == 'Отмена' or event == sg.WINDOW_CLOSED:
            break

        if event == 'add':
            month = values['month']
            base = values['base']
            file_name = values['file_name']

            if not file_name and base == 'Новая группа':
                sg.Popup('Введите название группы!', title='Ошибка')
                continue
            status = create_new_sheet(file_name, base, month)

            if not status:
                sg.Popup('Такая группа уже есть', title='Ошибка')
                continue
            if sg.Window("Добавить ещё",
                         [[sg.Text(f"Ведомость {status.name} добавлена. Добавить ещё?")], [sg.Yes(), sg.No()]],
                         element_justification='c', modal=True).read(close=True)[0] == "Yes":
                continue
            else:
                break

    window.close()


if __name__ == '__main__':
    pass
