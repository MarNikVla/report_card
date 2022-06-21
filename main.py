from itertools import chain, repeat
from openpyxl import load_workbook
from collections import Counter
import re

INITIAL_ROW_OF_NAMES = 13
FINAL_ROW_OF_NAMES = 49
COLUMN_OF_NAMES = 2
REPORT_CARD = 'test.xlsx'

wb = load_workbook(filename=REPORT_CARD)
machinist_sheet_name, DEM_sheet_name, reason_sheet_name = wb.sheetnames

machinist_sheet, DEM_sheet, reason_of_absence_sheet = wb[machinist_sheet_name], \
                                                      wb[DEM_sheet_name], \
                                                      wb[reason_sheet_name]


def get_lines_of_workers_dict(sheet) -> dict:
    lines_of_workers_dict = dict()

    for col in sheet.iter_cols(min_row=INITIAL_ROW_OF_NAMES,
                               max_row=FINAL_ROW_OF_NAMES,
                               min_col=COLUMN_OF_NAMES,
                               max_col=COLUMN_OF_NAMES):
        for cell in col:
            if cell.value is not None:
                lines_of_workers_dict[cell.value] = \
                    [cell.offset(row=i, column=j).value for i in [0, 1] for j in range(1, 17)]

    return lines_of_workers_dict


def normalize_lines_of_workers_dict(raw_dict: dict):
    normalized_dict = dict()
    for key, value in raw_dict.items():
        normalized_dict[key] = normalize_cells_list(value)
    return normalized_dict


def normalize_cells_list(cells_list: list[str]):
    new_cells_list = list()
    for cell in cells_list:
        if cell is not None:
            try:
                new_cell = float(cell.replace(",", "."))
            except ValueError:
                new_cell = re.sub(r'\s+', ' ', cell)
            new_cells_list.append(new_cell)

    new_cells_list.remove('Х')
    return new_cells_list


def count_days(lst: list):
    counter = Counter(lst)
    print(counter.keys())
    print(counter)
    absence_days = counter['В']
    vacation_days = counter['ОТ']
    medical_days = counter['Б']

    other_absence_days = sum(
        [counter[key] for key in ['ОВ', 'У', 'ДО', 'К', 'ПР', 'Р', 'ОЖ', 'ОЗ', 'Г', 'НН', 'НБ']])

    attendance_days = sum(
        counter.values()) - absence_days - vacation_days - medical_days - other_absence_days

    return attendance_days, absence_days, vacation_days, medical_days, other_absence_days


def count_hours(lst: list):
    hours = 0
    night_hours = 0
    counter = Counter(lst)
    print(counter)
    for i in counter.keys():
        if isinstance(i, float):
            hours += i * counter[i]
        elif i == '8/20':
            hours += 12 * counter[i]
        elif i in ['20/', '20/24']:
            hours += 4 * counter[i]
            night_hours += 2 * counter[i]
        elif i in ['/8 20/24', '0/8 20/', '/8 20/']:
            hours += 12 * counter[i]
            night_hours += 8 * counter[i]
        elif i in ['0/8', '/8']:
            hours += 8 * counter[i]
            night_hours += 6 * counter[i]

    print(hours)
    print(night_hours)
    return hours, night_hours

def write_cells():
    pass
if __name__ == '__main__':
    # intenize_from_list(line_of_workers['Буржинский А.В. Эл.монтер ЩУ ГТУ'])
    # intenize_from_list(lines_of_workers['Канева М.А. Уборщица '])
    # cont_attendance_days(lines_of_workers['Канева М.А. Уборщица '])
    # cont_attendance_days(lines_of_workers['Баранов Р.А. Эл.монтер ЩУ ГТУ'])
    print(count_days(
        normalize_cells_list(
            get_lines_of_workers_dict(DEM_sheet)['Баранов Р.А. Эл.монтер ЩУ ГТУ'])))
    print(count_hours(
        normalize_cells_list(
            get_lines_of_workers_dict(DEM_sheet)['Баранов Р.А. Эл.монтер ЩУ ГТУ'])))
