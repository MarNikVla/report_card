from functools import lru_cache
from itertools import compress
from typing import Type

from openpyxl import load_workbook
from collections import Counter
import re

from openpyxl.worksheet.worksheet import Worksheet

REPORT_CARD_FILE = 'test.xlsx'
BACKUP_REPORT_CARD_FILE = f'backup_{REPORT_CARD_FILE}'

wb = load_workbook(filename=REPORT_CARD_FILE)
machinist_sheet_name, DEM_sheet_name, reason_sheet_name = wb.sheetnames

machinist_sheet, DEM_sheet, reason_of_absence_sheet = wb[machinist_sheet_name], \
                                                      wb[DEM_sheet_name], \
                                                      wb[reason_sheet_name]


class Worker:
    def __init__(self, cell_index, sheet):
        self.cell_index = cell_index
        self.sheet = sheet

