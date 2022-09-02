import win32com.client
import os


def read_xls(file_name):
    path = f'{os.path.abspath(file_name)}.xls'
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path)
    sheet = wb.ActiveSheet

    i = 16
    kids = [sheet.Cells(i, 2).value]
    while sheet.Cells(i, 2).value:
        i += 1
        kids.append(sheet.Cells(i, 2).value)

    wb.Close()
    Excel.Quit()
    return kids[:-1]

def write_xls(data, file_name, day=1, row=16):
    Excel = win32com.client.Dispatch("Excel.Application")
    path = f'{os.path.abspath(file_name)}.xls'
    wb = Excel.Workbooks.Open(path)
    sheet = wb.ActiveSheet

    col = day + 4

    if len(data) == 1:
        col = 2
        row += 16


    for kid in data:
        sheet.Cells(row, col).value = kid
        row += 1

    wb.Save()
    wb.Close()
    Excel.Quit()

def check_directory():
    if not os.path.exists('Ведомости'):
        os.mkdir('Ведомости')


def main():
   check_directory()


if __name__ == '__main__':
    main()
