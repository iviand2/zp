import pandas as pd
from PySimpleGUI import PopupGetFile, PopupGetText, PopupError
from openpyxl import load_workbook


def get_file():
    return PopupGetFile('Пожалуйста укажите файл с трудовым календарем')


if __name__ == '__main__':
    file = get_file()
    months = ['январь',
              'февраль',
              'март',
              'апрель',
              'май',
              'июнь',
              'июль',
              'август',
              'сентябрь',
              'октябрь',
              'ноябрь',
              'декабрь']
    work_days = {}
    # book = pd.ExcelFile(file)
    books = load_workbook(file, data_only=True)
    try:
        year = int(PopupGetText('Пожалуйста, введите год загружаемого файла графика.\nФормат - 4 числовых символа,'
                                ' например 2020'))
    except:
        PopupError('Не смогли распознать введенный год, проверьте корректность формата!')
        raise ValueError('Не смогли распознать введенный год, проверьте корректность формата!')
    try:
        for ws in books.worksheets:
            sheet = ws.title
            counter = -1
            start = -1
            end = -1
            month_read = False
            data = []
            plan_data = []
            colors = []
            plan_colors = []
            for row in ws.rows:
                if 'график' in [c.value.lower() for c in row if type(c.value) is str]:
                    for cell_index, cell in enumerate(row):
                        if month_read:
                            month_num = months.index(cell.value.lower()) + 1
                            month = f'0{month_num}' if month_num < 10 else f'{month_num}'
                            month_read = False
                        elif 'график' in str(cell.value).lower():
                            start = cell_index
                            counter = 0
                            month_read = True
                        elif 'итог' in str(cell.value).lower():
                            end = cell_index
                            break
                if counter < 0:
                    continue
                elif counter == 2:
                    for cell in row[start+1: end]:
                        plan_data.append(cell.value)
                        colored = False
                        if cell.fill.fgColor.type == "rgb" and cell.fill.fgColor.rgb != '00000000':
                            colored = True
                        elif cell.fill.fgColor.type == 'indexed' and cell.fill.fgColor.index not in [1, 64, 63]:
                            colored = True
                        elif cell.fill.fgColor.type == 'theme' and cell.fill.fgColor.tint != 0 or cell.fill.fgColor.theme != 0:
                            colored = True
                        elif cell.fill.bgColor.type == "rgb" and cell.fill.fgColor.rgb != '00000000':
                            colored = True
                        elif cell.fill.bgColor.type == 'indexed' and cell.fill.bgColor.index not in [1, 64, 63]:
                            colored = True
                        elif cell.fill.bgColor.type == 'theme' and cell.fill.bgColor.tint != 0 or cell.fill.bgColor.theme != 0:
                            colored = True
                        plan_colors.append(colored)
                    counter += 1
                elif counter == 3:
                    for cell in row[start+1: end]:
                        data.append(cell.value)
                        colored = False
                        if cell.fill.fgColor.type == "rgb" and cell.fill.fgColor.rgb != '00000000':
                            colored = True
                        elif cell.fill.fgColor.type == 'indexed' and cell.fill.fgColor.index not in [1, 64, 63]:
                            colored = True
                        elif cell.fill.fgColor.type == 'theme' and cell.fill.fgColor.tint != 0:
                            colored = True
                        elif cell.fill.bgColor.type == "rgb" and cell.fill.bgColor.rgb != '00000000':
                            colored = True
                        elif cell.fill.bgColor.type == 'indexed' and cell.fill.bgColor.index not in [1, 64, 63]:
                            colored = True
                        elif cell.fill.bgColor.type == 'theme' and cell.fill.bgColor.tint != 0:
                            colored = True
                        colors.append(colored)
                    break
                else:
                    counter += 1
            if data:
                work_days[sheet] = {'Отработано': []}
                for index, day in enumerate(data):
                    worked_day = False
                    try:
                        worked_day = True if day is not None and float(day) else False
                    except ValueError:
                        pass
                    if worked_day:
                        day = index + 1
                        real_day = f'0{day}' if day < 10 else f'{day}'
                        appendix = ' Внеплановый' if colors[index] \
                                                     or (plan_data[index] is not None
                                                         and not str(plan_data[index]).isdigit()) \
                                                     or plan_colors[index] \
                            else ' Плановый'
                        work_days[sheet]['Отработано'].append(f'{year}.{month}.{real_day}{appendix}')

        ww = {c[0]: c[1]['Отработано'] for c in work_days.items()}
        pd.DataFrame.from_dict(ww, orient='index').dropna(how='all').T.to_excel('result.xlsx')
    except Exception as Ex:
        PopupError(str(Ex))

