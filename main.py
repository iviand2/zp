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
            # try:
                sheet = ws.title
                # data = book.parse(sheet_name=sheet)#.dropna(how='all', axis=1)
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
                        # if sheet.lower() == 'кобзев':
                        #     a = 1 + 1
                        for cell in row[start+1: end]:
                            plan_data.append(cell.value)
                            colored = False
                            if str(cell.fill.fgColor.rgb) != "Values must be of type <class 'str'>":
                                if cell.fill.fgColor.rgb != '00000000':
                                    colored = True
                            if cell.fill.fgColor.tint != 0:
                                colored = True
                            if cell.fill.bgColor.index != '00000000':
                                colored = True
                            plan_colors.append(colored)
                        counter += 1
                    elif counter == 3:
                        for cell in row[start+1: end]:
                            data.append(cell.value)
                            colored = False
                            if str(cell.fill.fgColor.rgb) != "Values must be of type <class 'str'>":
                                if cell.fill.fgColor.rgb != '00000000':
                                    colored = True
                            if cell.fill.fgColor.tint != 0:
                                colored = True
                            if cell.fill.bgColor.index != '00000000':
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

                        # data = data[data.columns[index:]]
                        # col_index = index
                        # break
                # for index, col in enumerate(data.columns):
                #     for cell in data[col]:
                #         if 'итог' in str(cell).lower():
                #             end = index
                # skipped = True
                # true_data = []
                # for index, row in enumerate(data.fillna(0).values):
                #     if row[0] and row[0].lower() == 'график':
                #         skipped = False
                #         month_num = months.index(row[1]) + 1
                #         month = f'0{month_num}' if month_num < 10 else f'{month_num}'
                #         row_index = index
                #     if not skipped:
                #         true_data.append(row)
                # days = true_data[1]
                # plan_work = true_data[2]
                # worked = true_data[3]
                # colors = []
                # for row_num, row in enumerate(books[sheet].rows):
                #     if row_num != row_index:
                #         continue
                #     else:
                #         for cell_num, cell in enumerate(row):
                #             if cell_num >= col_index:
                #                 colors.append(cell.fill.fgColor.rgb)
                #
                # for index, work in enumerate(worked[1:end]):
                #     if work and type(work) in [float, int]:
                #         appendix = ' Н' if plan_work[1:end][index] else ' П'
                #         day = index + 1
                #         real_day = f'0{day}' if day < 10 else f'{day}'
                #         work_days[sheet]['Отработано'].append(f'{year}.{month}.{real_day}{appendix}')
            # except IndexError:
            #     raise
            #     continue
        ww = {c[0]: c[1]['Отработано'] for c in work_days.items()}
        pd.DataFrame.from_dict(ww, orient='index').dropna(how='all').T.to_excel('result.xlsx')
    except Exception as Ex:
        raise Ex
        PopupError(str(Ex))

