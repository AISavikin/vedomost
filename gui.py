import PySimpleGUI as sg
import locale
from datetime import datetime
from func import *

locale.setlocale(locale.LC_ALL, 'ru-RU')
date = datetime.now()

sg.theme('DarkTeal10')

def main_window():
    # Основное окно приложения, само по себе ничего не делает, по сути навигационное меню.
    # Единственное значение которое можно передать дальше название группы

    # Проверяем существование директории "Ведомости" если нет, создаём
    check_directory()

    # Собираем файлы ведомостей, для добавления в выпадающий список, если файлов нет, создается пустой список
    try:
        list_group = [group for group in os.listdir(path=r'Ведомости/') if '.xlsx' in group]
        default_val = list_group[0]
    except IndexError:
        list_group = []
        default_val = ''

    layout = [
        [sg.Text("Ведомости детский сад")],
        [sg.Button('Добавить группу', expand_x=True), sg.Button('Добавить ученика')],
        [sg.Combo(list_group, expand_x=True, default_value=default_val, key='file_name', readonly=True),
         sg.Button('Отметить')],
        [sg.Button('test')]
    ]

    window = sg.Window('Ведомости', layout, element_justification='center')
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break

        if event == 'Добавить ученика':
            add_new_kids(values['file_name'])

        if event == 'Добавить группу':
            add_new_sheet(values['file_name'], list_group)

        if event == 'Отметить':
            check_kids(values['file_name'])

        if event == 'test':
            test(values['file_name'])

    window.close()


def add_new_kids(file_name: str):
    """
    Создает окно для добавления новых детей в ведомость.

    :param file_name: str: имя файла без расширения
    """
    # Получаем имена и фамилии детей в виде списка строк
    kids = get_kids(file_name)

    # Левая колонка
    left_col = [
        [sg.Text('Фамилия и имя')],
        [sg.Input(key='name', focus=True, size=25)],
        [sg.Button('Добавить'), sg.Button('Сохранить')]
    ]
    # Правая колонка
    right_col = [
        [sg.Text(file_name)],
        [sg.Table(headings=['№', 'Имя'], values=list(enumerate(kids, 1)), key='table', col_widths=[7, 20],
                  auto_size_columns=False, justification='center', size=(20, 15))]
    ]

    layout = [[sg.Column(left_col), sg.Column(right_col, element_justification='center')]]
    window = sg.Window(f'Добавить ученика в группу {file_name}', layout, modal=True)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break

        if event == 'Добавить' and values['name']:
            kids.append(values['name'])
            window['name'].update('')
            window['table'].update(values=enumerate(kids, 1))

        if event == 'Сохранить':
            save_new_kids(file_name, kids)
            break

    window.close()


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
        [sg.Combo(['Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май'],
                  size=(18, 0), key='month', default_value=f'{date:%B}', readonly=True)]
    ]
    layout = [
        [sg.Column(left), sg.Column(right)],
        [sg.Button('Добавить ведомость', key='add'), sg.Button('Отмена')]
    ]

    window = sg.Window('Добавить новую ведомость', layout, element_justification='center', modal=True)

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

            create_new_sheet(file_name, base, month)
            break

    window.close()


def check_kids(file_name: str):
    """
    Создает окно для проверки отсутствующих
    :param file_name: str имя файла без расширения
    """
    # Получаем список детей из файла Excel
    kids = get_kids(file_name)
    # Левая колонка: текстовые виджеты с именами детей из списка kids
    left_col = [[sg.Text(kid)] for kid in kids]
    # Правая колонка с пустыми инпутами, пронумерованные от нуля
    right_col = [[sg.Input(size=(2, 1), key=i)] for i in range(len(kids))]
    # Устанавливаем фокус на первого ребенка
    right_col[0][0].Focus = True

    layout = [
        [sg.Frame('Дата',
                  [[sg.Text('Число: '),
                    sg.Combo(list(range(1, 32)), default_value=f'{date:%d}', k='date'),
                    sg.Text(f'{date:%B}')]])],
        [sg.Text(f'{file_name}')],
        [sg.Column(left_col), sg.Column(right_col)],
        [sg.Button('Отметить')]
    ]

    window = sg.Window('Отметить ученика', layout, finalize=True, element_justification='center',
                       return_keyboard_events=True)
    # При фокусе на кнопке с ключом "Отметить", нажатие Enter генерирует событие "Отметить_Enter"
    window['Отметить'].bind("<Return>", "_Enter")

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED:
            break
        focus = window.find_element_with_focus()
        prev_focus = focus.get_previous_focus()
        next_focus = focus.get_next_focus()

        if event == 'Right:39' and type(focus) == sg.PySimpleGUI.Input:
            focus.update('б')

        if event == 'Left:37' and type(focus) == sg.PySimpleGUI.Input:
            focus.update('н')

        if event == 'Up:38':
            prev_focus.set_focus()

        if event == 'Down:40':
            next_focus.set_focus()

        if event == 'Отметить' or event == 'Отметить_Enter':
            absent = [values[i] for i in range(len(kids))]
            if not all(absent):
                sg.Popup('Вы отметили не всех!')
            else:
                check_absent(file_name, values['date'], absent)

    window.close()


def test(file_name):
    kids = ['Боровиков Яша', 'Всецина Нина', 'Кракозябра', 'Литвиненко Полина', 'Мушникова Полина', 'Саечкин Демид',
            'Донцова Олеся', 'Мазнева Валерия', 'Малехина Ярослава', 'Пиримова Амира', 'Тимашевская Анастасия',
            'Чугунов Ярослав', 'Бобрышева Маша', 'Дорошенко Алиса', 'Бегенова София']

    layout = [
        [sg.Text('qweqwe')],
        [sg.Input()],
        [sg.Button('OK')]
    ]

    window = sg.Window('Отметить ученика', layout, finalize=True, element_justification='center',
                       return_keyboard_events=True)

    while True:
        event, values = window.read()
        print(event)
        if event == sg.WINDOW_CLOSED:
            break

    window.close()


def main():
    main_window()


if __name__ == '__main__':
    main()
