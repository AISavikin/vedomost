import PySimpleGUI as sg
from calendar import Calendar
from func import MONTH_DICT, read_notes, save_notes
from openpyxl import load_workbook

lorem = """Lorem ipsum dolor  auctor, vestibulum lacinia velit. Sed dapibus vel dolor et semper.
"""


def notes_window(file_name: str):
    notes = read_notes(file_name)
    l_notes = {day: notes[day] for num, day in enumerate(notes, 1) if num % 2 == 1}
    r_notes = {day: notes[day] for num, day in enumerate(notes, 1) if num % 2 == 0}

    l_col = [[sg.Text(day),
              sg.Multiline(default_text=l_notes[day], size=(50, 6), key=day, no_scrollbar=True)] for day in l_notes]

    r_col = [[sg.Text(day),
              sg.Multiline(default_text=r_notes[day], size=(50, 6), key=day, no_scrollbar=True)] for day in r_notes]

    layout = [
        [sg.Text(file_name, font='Times_new_roman 21 bold'), sg.Button('Сохранить', font='Times_new_roman 17')],
        [sg.Column(l_col, scrollable=False, size_subsample_width=1, element_justification='r', size_subsample_height=1),
         sg.Column(r_col, scrollable=False, size_subsample_width=1, element_justification='r', size_subsample_height=1),
         ],
    ]

    window = sg.Window('Заметки', layout, element_justification='center')

    while True:
        event, values = window.read()

        if event == 'Отмена' or event == sg.WINDOW_CLOSED:
            break

        if event == 'Сохранить':
            if sg.Window("Добавить ещё", [[sg.Text("Ведомость !!! добавлена. Добавить ещё?")],
                                          [sg.Yes(), sg.No()]], element_justification='c').read(close=True)[0] == "Yes":
                continue
            else:
                break
            # print(values)
            # notes = {day: values[day] for day in sorted(values)}
            # print(notes)
            # notes = {key: notes[key] for key in sorted(notes)}
            # save_notes(file_name, notes)
            # break
    window.close()

def main():
    notes_window('Группа 1_Октябрь.xlsx')
    notes = read_notes('Группа 1_Октябрь.xlsx')
    # l = {}
    # r = {}
    # for num, day in enumerate(notes, 1):
    #     if num % 2 == 1:
    #         l[day] = notes[day]
    #     else:
    #         r[day] = notes[day]
    #
    l = {day: notes[day] for num, day in enumerate(notes, 1) if num % 2 == 1}
    r = {day: notes[day] for num, day in enumerate(notes, 1) if num % 2 == 0}
    print(l)
    key_l = sorted(l)
    l = {key: l[key] for key in key_l}
    print(l)
    print(r)


if __name__ == '__main__':
    main()
