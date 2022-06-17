from openpyxl import load_workbook
from collections import  Counter
from random import randint

wb = load_workbook(filename='test.xlsx')
machinist_sheet_name, DEM_sheet_name, reason_sheet_name = wb.sheetnames

machinist_sheet, DEM_sheet, reason_of_absence_sheet = wb[machinist_sheet_name], \
                                                      wb[DEM_sheet_name], \
                                                      wb[reason_sheet_name]
line_of_workers = dict()

# for i in machinist_sheet['b24:b25']:
    # print(i[0].value)

for col in DEM_sheet.iter_cols(min_row=13, max_row=49, min_col=2, max_col=2):
    for cell in col:
        if cell.value is not None:
            # print(cell)
            line_of_workers[cell.value] = [cell.offset(row=i, column=j).value for i in [0, 1] for j
                                           in range(1, 17)]

attendance_days = 0
absence_days = 0
vacation_days = 0
medical_days = 0
other_absence_days = 0
hours = 0
night_hours = 0

def cont_attendance_days(lst:list):
    c= Counter(lst)
    print(c)



def intenize_from_list(lst: list):
    new_lst = list()
    for string in lst:
        try:
            new_lst.append(float(string.replace(",", ".")))
        except ValueError:
            new_lst.append(string)
    print(new_lst)
    return new_lst

attend_days = 0
absence_days = 0
vacation_days = 0
medical_days = 0
other_absence_days = 0
hours = 0
night_hours = 0

def cont_days(intenize_list:list):
    for item in intenize_list:
        attend_days = item.count
    pass




if __name__ == '__main__':
    # intenize_from_list(line_of_workers['Буржинский А.В. Эл.монтер ЩУ ГТУ'])
    intenize_from_list(line_of_workers['Канева М.А. Уборщица '])
    cont_attendance_days(line_of_workers['Канева М.А. Уборщица '])
    cont_attendance_days(line_of_workers['Буржинский А.В. Эл.монтер ЩУ ГТУ'])
