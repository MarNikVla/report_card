
from functools import lru_cache
import pathlib
from typing import Type

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from report_сard_via_classes import Worker

REPORT_CARD_FILE = pathlib.Path('C:/project/report_card/test.xlsx')

# def test(file_name):
#     print(pathlib.Path.home())
#     print(file_name.name)
#     print(file_name.parent.joinpath(f'backup_{file_name.name}'))
#     print(file_name.is_file())
# with file_name.open() as f:
#     f.read()


# REPORT_CARD_FILE = 'test.xlsx'
# print(REPORT_CARD_FILE)
# BACKUP_REPORT_CARD_FILE = REPORT_CARD_FILE.parent.joinpath(f'backup_{REPORT_CARD_FILE.name}')
INITIAL_ROW_OF_NAMES: int = 13
FINAL_ROW_OF_NAMES: int = 49
COLUMN_OF_NAMES: int = 2


# wb = load_workbook(filename=REPORT_CARD_FILE)
# machinist_sheet_name, DEM_sheet_name, reason_sheet_name = wb.sheetnames
#
# machinist_sheet, DEM_sheet, reason_of_absence_sheet = wb[machinist_sheet_name], \
#                                                       wb[DEM_sheet_name], \
#                                                       wb[reason_sheet_name]


@lru_cache
def get_workers(sheet: Type[Worksheet]) -> list:
    wokers_list = list()

    # Iteration on column of names
    for col in sheet.iter_cols(min_row=INITIAL_ROW_OF_NAMES,
                               max_row=FINAL_ROW_OF_NAMES,
                               min_col=COLUMN_OF_NAMES,
                               max_col=COLUMN_OF_NAMES):
        for cell in col:
            if cell.value is not None:
                wokers_list.append(Worker(cell.coordinate, sheet))
    # print(wokers_list)
    return wokers_list


def fill_all_workers(wokers_list: list):
    for worker in wokers_list:
        worker.fill_worker_line()


# def save_sheet(sheet):
#     BACKUP_REPORT_CARD_FILE = REPORT_CARD_FILE.parent.joinpath(f'backup_{REPORT_CARD_FILE.name}')
#     fill_all_workers(get_workers(sheet))
#     wb.save(BACKUP_REPORT_CARD_FILE)


def save_file(file_name):
    report_card_file = pathlib.Path(file_name)
    backup_report_card_file = report_card_file.parent.joinpath(f'backup_{report_card_file.name}')
    wb = load_workbook(filename=report_card_file)
    # print(wb._sheets)
    for sheet in wb._sheets:
        fill_all_workers(get_workers(sheet))
    wb.save(backup_report_card_file)


if __name__ == '__main__':
    # save_sheet(DEM_sheet)
    # save_sheet(machinist_sheet)
    # test(REPORT_CARD_FILE)
    save_file('C:/project/report_card/test.xlsx')
