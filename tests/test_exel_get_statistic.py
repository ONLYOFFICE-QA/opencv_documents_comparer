# import openpyxl
#
path = r'C:\Python\CompareIMG\tests\1.xls'
path2 = r'C:\Python\CompareIMG\tests\1.xlsx'
#
# wb = openpyxl.load_workbook(path)
# sheet = wb.worksheets[0]
#
# row_count = sheet.max_row
# column_count = sheet.max_column
# print(row_count)
# print(column_count)

import xlrd

rb = xlrd.open_workbook(path, formatting_info=True)
sheet = rb.sheet_by_index(1)
w = sheet.computed_column_width(0)

for i, row in sheet.rowinfo_map.items():
    print("row {0} has {1} twip height".format(i, row.height))

for i, col in sheet.colinfo_map.items():
    print("column {0} has {1} width".format(i, col.width))
