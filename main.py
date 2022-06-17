from openpyxl import load_workbook
from random import randint

wb = load_workbook(filename='test.xlsx')
machinist_sheet_name, DEM_sheet_name, reason_sheet_name = wb.sheetnames

machinist_sheet, DEM_sheet, reason_of_absence_sheet = wb[machinist_sheet_name], \
                                                      wb[DEM_sheet_name], \
                                                      wb[reason_sheet_name]
line_of_workers = dict()

for i in machinist_sheet['b24:b25']:
    print(i[0].value)

for col in DEM_sheet.iter_cols(min_row=13, max_row=49, min_col=2, max_col=2):
    for cell in col:
        if cell.value is not None:
            # print(cell)
            line_of_workers[cell.value] = [cell.offset(row=i, column=j).value for i in [0, 1] for j
                                           in range(1, 17)]

print(*line_of_workers.items(), sep='\n')
