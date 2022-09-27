import PySimpleGUI as sg

from func import *

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
        [sg.Input(key='name', focus=True, size=26)],
        [sg.Button('Сохранить'), sg.Button('Исправить'), sg.Button('+')],
        # [sg.Button('Удалить')]
    ]
    # Правая колонка
    right_col = [
        [sg.Text(file_name)],
        [sg.Table(headings=['№', 'Имя'], values=list(enumerate(kids, 1)), key='table', col_widths=[7, 20],
                  auto_size_columns=False, justification='center', size=(20, 15))]
    ]

    layout = [[sg.Column(left_col), sg.Column(right_col, element_justification='center')]]
    window = sg.Window(f'Добавить ученика в группу {file_name}', layout, modal=True, finalize=True)
    window.bind("<Return>", "+")

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break

        if event == '+' and values['name']:
            kids.append(values['name'])
            window['name'].update('')
            window['table'].update(values=enumerate(kids, 1))

        if event == 'Исправить' and values['table']:
            kids[values['table'][0]] = values['name']
            window['table'].update(values=enumerate(kids, 1))
            window['name'].update('')

        # if event == 'Удалить':
        #     kids.pop(values['table'][0])
        #     window['table'].update(values=enumerate(kids, 1))

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

            if not create_new_sheet(file_name, base, month):
                sg.Popup('Такая группа уже есть')
            break

    window.close()


def check_kids(file_name: str):
    """
    Создает окно для проверки отсутствующих
    :param file_name: str имя файла без расширения
    """
    month = file_name.split('_')[-1][:-5]
    work_days = get_work_days(MONTH_DICT[month])

    # Получаем список детей из файла Excel
    kids = get_kids(file_name)
    # Левая колонка: текстовые виджеты с именами детей из списка kids
    left_col = [[sg.Text(kid)] for kid in kids]
    # Правая колонка с пустыми инпутами, пронумерованные от нуля
    right_col = [[sg.Input(size=(2, 1), key=i)] for i in range(len(kids))]
    # Устанавливаем фокус на первого ребенка
    right_col[0][0].Focus = True

    layout = [
        [sg.Text(f'{file_name}')],
        [sg.Frame('Дата',
                  [[sg.Text('Число: '),
                    sg.Combo(work_days, default_value=f'{date:%d}', k='date'),
                    sg.Text(month)]])],
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
                mark_absent(file_name, values['date'], absent)
                break

    window.close()


def notes_window(file_name: str):
    notes = read_notes(file_name)

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

def main():
    pass


if __name__ == '__main__':
    main()
