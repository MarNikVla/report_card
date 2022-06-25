from copy import deepcopy

from openpyxl import load_workbook

REPORT_CARD_FILE = 'test.xlsx'
BACKUP_REPORT_CARD_FILE = f'backup_{REPORT_CARD_FILE}'

wb = load_workbook(filename=REPORT_CARD_FILE)
machinist_sheet_name, DEM_sheet_name, reason_sheet_name = wb.sheetnames

machinist_sheet, DEM_sheet, reason_of_absence_sheet = wb[machinist_sheet_name], \
                                                      wb[DEM_sheet_name], \
                                                      wb[reason_sheet_name]


class Sheet:
    def __init__(self, sheet):
        self.sheet = sheet

    @staticmethod
    def _type_of_day(cell):

        working_day = 'РД'  # Рабочий день
        weekend = 'В'  # Выходной
        short_day = 'КД'  # Короткий день
        holiday = 'П'  # Праздник

        color_standard_red = 'FFFF0000'
        color_standard_light_green = 'FF92D050'
        color_standard_orange = 'FFFFC000'

        if cell.fill.start_color.index == color_standard_red:
            return holiday
        if cell.fill.start_color.index == color_standard_light_green:
            return weekend
        if cell.fill.start_color.index == color_standard_orange:
            return short_day
        return working_day

    @property
    def work_days_matrix(self):
        first_day_of_month = self.sheet['C10']
        last_day_of_month = self.sheet['R11']
        work_days_matrix = list()
        for row in self.sheet.iter_rows(min_row=first_day_of_month.row,
                                        max_row=last_day_of_month.row,
                                        min_col=first_day_of_month.column,
                                        max_col=last_day_of_month.column):
            for cell in row:
                if cell.value not in (None, 'Х'):
                    work_days_matrix.append(self._type_of_day(cell))
        return work_days_matrix

    def __str__(self):
        return f'{self.sheet}'


class Worker(Sheet):
    def __init__(self, cell_index, sheet):
        super().__init__(sheet)
        self.cell = sheet[cell_index]

    def __str__(self):
        return f'{self.cell.value},row  {self.cell.row}, col {self.cell.column}'

    def __len__(self):
        return len(self.cells_range)

    @property
    def cells_range(self):
        cells_range = list()
        for col in self.sheet.iter_rows(min_row=self.cell.row,
                                        max_row=self.cell.row + 1,
                                        min_col=self.cell.column + 1,
                                        max_col=self.cell.column + 16):
            for cell in col:
                if cell.value not in (None, 'Х'):
                    cells_range.append(cell.value)
        return cells_range

    def is_28_hours(self):
        color_standard_orange = 'FFFFC000'
        return self.cell.fill.start_color.index == color_standard_orange

    @property
    def work_days_matrix(self):
        return super(Worker, self).work_days_matrix[:len(self.cells_range)]

    def normalize_workdays(self):
        days_to_remove = ['ОТ','У','ДО','Б','К','Р','ОЖ','ОЗ','Г','НН','НБ']
        normalize_workdays= deepcopy(self.work_days_matrix)
        for index, cell in enumerate(self.cells_range):
            if cell in days_to_remove:
                normalize_workdays[index] = 0
            normalize_workdays
        return normalize_workdays



worker = Worker('B13', DEM_sheet)
worker_women = Worker('B35', DEM_sheet)

print(worker)
print(worker.cells_range)
print(len(worker))
print(worker.work_days_matrix)
print(worker_women.is_28_hours())
print(worker.normalize_workdays())
