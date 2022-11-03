from conf import *
import PySimpleGUI as sg
from database import Attendance
from utils import get_kids, get_work_days, date, MONTH_DICT
from loguru import logger


def get_absents(kids, day, month):
    absent = []
    for kid in kids:
        for i in Attendance.filter(id=kid.id, day=day, month=month):
            if not i.absent:
                absent.append('')
                continue
            absent.append(i.absent)
    if not absent:
        absent = ['' for _ in range(len(kids))]
    return absent


def mark_absent(kids, absents, day, month_name):
    for kid in range(len(kids)):
        Attendance.create(absent=absents[kid], day=day, month=MONTH_DICT[month_name][0],
                          year=MONTH_DICT[month_name][1], id=kids[kid].id)


def mark_kids(num_group: int, month_name: str):
    """
    Создает окно для проверки отсутствующих
    :param month_name: str Название месяца
    :param num_group: int номер группы
    """
    # Получаем список рабочих дней
    work_days = get_work_days(month_name)
    # Получаем список детей из базы данных
    kids = get_kids(num_group)

    absents = get_absents(kids, date.day, MONTH_DICT[month_name][0])

    # Левая колонка: текстовые виджеты с именами детей из списка kids
    left_col = [[sg.Text(kid.name)] for kid in kids]
    # Правая колонка с пустыми инпутами, пронумерованные от нуля
    right_col = [[sg.Input(size=(2, 1), key=i, default_text=absents[i])] for i in range(len(kids))]
    # Устанавливаем фокус на первого ребенка
    right_col[0][0].Focus = True

    layout = [
        [sg.Text(f'Группа {num_group}')],
        [sg.Frame('Дата',
                  [[sg.Text('Число: '),
                    sg.Combo(work_days, default_value=date.day, k='date', enable_events=True),
                    sg.Text(month_name)]])],
        [sg.Column(left_col), sg.Column(right_col)],
        [sg.Button('Отметить')]
    ]

    window = sg.Window('Отметить ученика', layout, finalize=True, element_justification='center',
                       return_keyboard_events=True, font=(FONT_FAMILY, FONT_SIZE))
    # При фокусе на кнопке с ключом "Отметить", нажатие Enter генерирует событие "Отметить_Enter"
    window['Отметить'].bind("<Return>", "_Enter")

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        focus = window.find_element_with_focus()
        prev_focus = focus.get_previous_focus()
        next_focus = focus.get_next_focus()
        if event == 'date':
            absent = get_absents(kids, values['date'], MONTH_DICT[month_name][0])
            for i in range(len(absent)):
                window[i].Update(absent[i])

        if event == 'Right:39' and type(focus) is sg.PySimpleGUI.Input:
            focus.update('б')

        if event == 'Left:37' and type(focus) is sg.PySimpleGUI.Input:
            focus.update('н')

        if event == 'Up:38':
            prev_focus.set_focus()

        if event == 'Down:40':
            if type(focus) is sg.PySimpleGUI.Button:
                pass
            else:
                next_focus.set_focus()

        if event == 'Отметить' or event == 'Отметить_Enter':
            absents = [values[i] for i in range(len(kids))]
            logger.info({k.name: v for k in kids for v in absents})
            if not all(absents):
                if sg.Window('Вы уверены?', [
                    [sg.Text('Вы отметили не всех! Сохранить?'), sg.Button('Да'),
                     sg.Button('Нет')]]).read(close=True)[0] == 'Нет':
                    continue
            mark_absent(kids, absents, values['date'], month_name)
            break

    window.close()

if __name__ == '__main__':
    mark_kids(7, 'Октябрь')