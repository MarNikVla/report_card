from functools import lru_cache
from typing import Type

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from report_сard_via_classes import Worker

REPORT_CARD_FILE = 'C:/projects/report_card/test.xlsx'
BACKUP_REPORT_CARD_FILE = f'backup_{REPORT_CARD_FILE}'
INITIAL_ROW_OF_NAMES = 13
FINAL_ROW_OF_NAMES = 49
COLUMN_OF_NAMES = 2

wb = load_workbook(filename=REPORT_CARD_FILE)
machinist_sheet_name, DEM_sheet_name, reason_sheet_name = wb.sheetnames

machinist_sheet, DEM_sheet, reason_of_absence_sheet = wb[machinist_sheet_name], \
                                                      wb[DEM_sheet_name], \
                                                      wb[reason_sheet_name]


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


def save_sheet(sheet):
    fill_all_workers(get_workers(sheet))
    wb.save(BACKUP_REPORT_CARD_FILE)

def save_file(file_name):
    pass



if __name__ == '__main__':
    save_sheet(DEM_sheet)
    save_sheet(machinist_sheet)