import xlrd
import xlwt
from xlutils.copy import copy
import os

FIPS2GEO = {}
with open("counties.txt",encoding = "utf-8") as county:
    lines = county.readlines()
    for line in lines[1:]:
        line = line.split()
        FIPS2GEO[line[1]]=(line[9],line[10])

FIPS2GEO['51515']=(45.13,-72.96)

data = xlrd.open_workbook("MCM_NFLIS_Data.xls",formatting_info=True)
old_sheet = data.sheets()[1]
rows = old_sheet.nrows
cols = old_sheet.ncols
new_data = copy(data)
table = new_data.get_sheet(1)
table.write(0, 10, "LAT")
table.write(0, 11, "LONG")
for i in range(1,rows):
    FIPS_Combined = str(old_sheet.row_values(i)[5])
    table.write(i, 10, FIPS2GEO[FIPS_Combined][0])
    table.write(i, 11, FIPS2GEO[FIPS_Combined][1])
new_data.save("data_GEO.xls")

