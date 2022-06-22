from functools import lru_cache
from itertools import chain, repeat
from openpyxl import load_workbook
from collections import Counter
import re

INITIAL_ROW_OF_NAMES = 13
FINAL_ROW_OF_NAMES = 49
COLUMN_OF_NAMES = 2
REPORT_CARD_FILE = 'test.xlsx'

wb = load_workbook(filename=REPORT_CARD_FILE)
machinist_sheet_name, DEM_sheet_name, reason_sheet_name = wb.sheetnames

machinist_sheet, DEM_sheet, reason_of_absence_sheet = wb[machinist_sheet_name], \
                                                      wb[DEM_sheet_name], \
                                                      wb[reason_sheet_name]


@lru_cache
def get_lines_of_working_days(sheet) -> dict:
    lines_of_working_days = dict()

    for col in sheet.iter_cols(min_row=INITIAL_ROW_OF_NAMES,
                               max_row=FINAL_ROW_OF_NAMES,
                               min_col=COLUMN_OF_NAMES,
                               max_col=COLUMN_OF_NAMES):
        for cell in col:
            if cell.value is not None:
                lines_of_working_days[cell.coordinate] = \
                    [cell.offset(row=i, column=j).value for i in [0, 1] for j in range(1, 17)]

    return lines_of_working_days


def get_normalize_cells_list(cells_list: list[str]) -> list:
    new_cells_list = list()
    for cell in cells_list:
        if cell is not None:
            try:
                new_cell = float(cell.replace(",", "."))
            except ValueError:
                new_cell = re.sub(r'\s+', '', cell)
            new_cells_list.append(new_cell)

    new_cells_list.remove('Х')
    return new_cells_list


def count_days(lst: list) -> dict:
    counter = Counter(lst)
    days_dict = dict()

    absence_days = counter['В']
    absence_paid_days = counter['НОД']
    vacation_days = counter['ОТ']
    medical_days = counter['Б']
    other_absence_days = sum(
        [counter[key] for key in ['ОВ', 'У', 'ДО', 'К', 'ПР', 'Р', 'ОЖ', 'ОЗ', 'Г', 'НН', 'НБ']])
    attendance_days = sum(
        counter.values()) - (
                                  absence_days + vacation_days + medical_days +
                                  other_absence_days + absence_paid_days)

    days_dict['attendance_days'] = attendance_days or None
    days_dict['absence_days'] = absence_days or None
    days_dict['vacation_days'] = vacation_days or None
    days_dict['medical_days'] = medical_days or None
    days_dict['other_absence_days'] = other_absence_days or None
    days_dict['absence_paid_days'] = absence_paid_days or None

    return days_dict


def count_hours(lst: list) -> dict:
    hours = 0
    night_hours = 0
    counter = Counter(lst)
    hours_dict = dict()
    for i in counter.keys():
        if isinstance(i, float):
            hours += i * counter[i]
        elif i == '8/20':
            hours += 12 * counter[i]
        elif i in ['20/', '20/24']:
            hours += 4 * counter[i]
            night_hours += 2 * counter[i]
        elif i in ['/820/24', '0/820/', '/820/']:
            hours += 12 * counter[i]
            night_hours += 8 * counter[i]
        elif i == '820/':
            hours += 12 * counter[i]
            night_hours += 2 * counter[i]
        elif i == '420/':
            hours += 8 * counter[i]
            night_hours += 2 * counter[i]
        elif i in ['0/8', '/8']:
            hours += 8 * counter[i]
            night_hours += 6 * counter[i]

    hours_dict['hours'] = hours or None
    hours_dict['night_hours'] = night_hours or None

    return hours_dict


def write_days_to_file(sheet, cell_index: str):
    print(sheet[cell_index])
    print(sheet[cell_index].offset(row=0, column=17))
    print(sheet[cell_index].offset(row=0, column=29))
    lines_of_working_days = get_lines_of_working_days(sheet)[cell_index]
    normalize_cells_list = get_normalize_cells_list(lines_of_working_days)
    days = count_days(normalize_cells_list)

    # cell of attendance_days = cell_index.offset(row=0, column=17)
    sheet[cell_index].offset(row=0, column=17).value = days['attendance_days']

    # cell of absence_days = cell_index.offset(row=0, column=22)
    sheet[cell_index].offset(row=0, column=22).value = days['absence_days']

    # cell of vacation_days = cell_index.offset(row=0, column=23)
    sheet[cell_index].offset(row=0, column=23).value = days['vacation_days']

    # cell of medical_days = cell_index.offset(row=0, column=24)
    sheet[cell_index].offset(row=0, column=24).value = days['medical_days']

    # cell of other_absence_days = cell_index.offset(row=0, column=25)
    sheet[cell_index].offset(row=0, column=25).value = days['other_absence_days']

    # cell of absence_paid_days = cell_index.offset(row=0, column=26)
    sheet[cell_index].offset(row=0, column=26).value = days['absence_paid_days']

    # wb.save('document.xlsx')
    pass


def write_hours_to_file(sheet, cell_index: str):
    print(sheet[cell_index])
    print(sheet[cell_index].offset(row=0, column=17))
    print(sheet[cell_index].offset(row=0, column=29))
    lines_of_working_days = get_lines_of_working_days(sheet)[cell_index]
    normalize_cells_list = get_normalize_cells_list(lines_of_working_days)
    hours = count_hours(normalize_cells_list)

    # cell of hours = cell_index.offset(row=0, column=18)
    sheet[cell_index].offset(row=0, column=18).value = hours['hours']

    # cell of night_hours = cell_index.offset(row=0, column=22)
    sheet[cell_index].offset(row=0, column=20).value = hours['night_hours']

    # wb.save('document.xlsx')
    pass


def write_to_file(sheet):
    lines_of_working_days_dict = get_lines_of_working_days(sheet)
    for key in lines_of_working_days_dict.keys():
        write_days_to_file(sheet, key)
        write_hours_to_file(sheet, key)

    # write_days_to_file(sheet):
    wb.save('document.xlsx')
    pass

def get_holidays(sheet):
    # days_range = sheet['C10':'R11']
    # print(days_range)
    lines_of_days = dict()
    min_row = 10
    min_col = 3
    for col in sheet.iter_rows(min_row=min_row,
                               max_row=min_row+1,
                               min_col=min_col,
                               max_col=min_col+15):
        for cell in col:
            # print(cell.color)
            if cell.value is not None:
                lines_of_days[cell.coordinate] = cell.value


    return lines_of_days

    pass
# def get_lines_of_working_days(sheet) -> dict:
#     lines_of_working_days = dict()
#
#     for col in sheet.iter_cols(min_row=INITIAL_ROW_OF_NAMES,
#                                max_row=FINAL_ROW_OF_NAMES,
#                                min_col=COLUMN_OF_NAMES,
#                                max_col=COLUMN_OF_NAMES):
#         for cell in col:
#             if cell.value is not None:
#                 lines_of_working_days[cell.coordinate] = \
#                     [cell.offset(row=i, column=j).value for i in [0, 1] for j in range(1, 17)]
#
#     return lines_of_working_days



if __name__ == '__main__':
    # intenize_from_list(line_of_workers['Буржинский А.В. Эл.монтер ЩУ ГТУ'])
    # intenize_from_list(lines_of_workers['Канева М.А. Уборщица '])
    # cont_attendance_days(lines_of_workers['Канева М.А. Уборщица '])
    # cont_attendance_days(lines_of_workers['Баранов Р.А. Эл.монтер ЩУ ГТУ'])
    # print(count_days(
    #     normalize_cells_list(
    #         get_lines_of_working_days(DEM_sheet)['B13'])))
    # print(count_hours(
    #     normalize_cells_list(
    #         get_lines_of_working_days(DEM_sheet)[15])))
    # print(write_days_to_file(DEM_sheet, 'B13'))
    print(write_to_file(DEM_sheet))
    # print(write_hours_to_file(DEM_sheet, 'B13'))
    # print(get_holidays(DEM_sheet))
