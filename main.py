from itertools import chain, repeat
from openpyxl import load_workbook
from collections import Counter
import re

INITIAL_ROW_OF_NAMES = 13
FINAL_ROW_OF_NAMES = 49
COLUMN_OF_NAMES = 2

wb = load_workbook(filename='test.xlsx')
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
                # print(cell)
                lines_of_workers_dict[cell.value] = \
                    [cell.offset(row=i, column=j).value for i in [0, 1] for j in range(1, 17)]

    return lines_of_workers_dict


def normalize_lines_of_workers_dict(raw_dict: dict):
    # removing ('X' from raw_dict (pop(15))
    # lines_of_workers_dict = {k: v for (k, v) in filter(lambda x: x[1].pop(15), raw_dict.items())}
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
                # cell.strip()
                new_cell = re.sub(r'\s+', ' ', cell)
            new_cells_list.append(new_cell)

    return new_cells_list


attendance_days = 0
absence_days = 0
vacation_days = 0
medical_days = 0
other_absence_days = 0
hours = 0
night_hours = 0


def cont_attendance_days(lst: list):
    lst.remove('Х')
    c = Counter(lst)
    print(c)
    print(c)
    print(c.total())
    print(c['В'])



# def intenize_from_list(lst: list):
#     new_lst = list()
#     for string in lst:
#         try:
#             new_lst.append(float(string.replace(",", ".")))
#         except ValueError:
#             new_lst.append(string)
#     print(new_lst)
#     return new_lst


attend_days = 0


def cont_days(cells_list: list):
    for item in cells_list:
        attend_days = item.count

    pass


if __name__ == '__main__':
    # intenize_from_list(line_of_workers['Буржинский А.В. Эл.монтер ЩУ ГТУ'])
    # intenize_from_list(lines_of_workers['Канева М.А. Уборщица '])
    # cont_attendance_days(lines_of_workers['Канева М.А. Уборщица '])
    # cont_attendance_days(lines_of_workers['Буржинский А.В. Эл.монтер ЩУ ГТУ'])
    print(cont_attendance_days(
        normalize_cells_list(get_lines_of_workers_dict(DEM_sheet)['Буржинский А.В. Эл.монтер ЩУ ГТУ'])))
