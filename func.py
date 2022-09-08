import win32com.client
import os


def excel(func):
    def wrapper(file_name, *args, **kwargs):
        Excel = win32com.client.Dispatch("Excel.Application")
        path = f'{os.path.abspath(file_name)}.xlsx'
        wb = Excel.Workbooks.Open(path)
        sheet = wb.ActiveSheet

        res = func(sheet, *args, **kwargs)

        wb.Save()
        wb.Close()
        Excel.Quit()
        return res

    return wrapper


@excel
def read_xls(sheet):
    vals = []
    row = 16
    while sheet.Cells(row, 2).value:
        vals.append(sheet.Cells(row, 2).value)
        row += 1
    return vals


@excel
def write_absent(sheet, **kwargs):
    data, day = kwargs['data'], kwargs['day']
    col = day + 4
    row = 16
    for kid in data:
        sheet.Cells(row, col).value = kid
        row += 1


@excel
def write_new_kids(sheet, **kwargs):
    row, new_kids = kwargs['row'] + 15, kwargs['new_kids']
    col = 2
    for kid in new_kids:
        sheet.Cells(row, col).value = kid
        row += 1


@excel
def write_month(sheet, month, group_name):
    data = ['' for _ in range(22)]
    for col in range(5, 38):
        row = 16
        for rec in data:
            sheet.Cells(row, col).value = rec
            row += 1
    sheet.Cells(3, 14).value = month
    sheet.Cells(42, 27).value = month
    sheet.Cells(5, 3).value = group_name



def check_directory():
    if not os.path.exists('Ведомости'):
        os.mkdir('Ведомости')


def main():
    check_directory()
    kids = read_xls('Template — копия')
    print(kids)


if __name__ == '__main__':
    main()
