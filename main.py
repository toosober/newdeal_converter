import xlrd
# function should return date of last update of source file
def getvalidationdate():
    return '11//11/19'

rb = xlrd.open_workbook('c:/test_data.xls', formatting_info=True)
sheet = rb.sheet_by_index(0)
for rownum in range(sheet.nrows):
    row = sheet.row_values(rownum)
    for c_el in row:
        print(getvalidationdate())
        if c_el != '':
            print(c_el)



