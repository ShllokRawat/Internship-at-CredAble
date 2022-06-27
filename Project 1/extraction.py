import xlrd
import xlwt
from xlwt import Workbook
wb = Workbook()
sheet1 = wb.add_sheet('Extracted Data')
location = str("C:\\Users\\shllo\\Desktop\\Internship\\Data\\RMWH1649332131561.xlsx")
wb2 = xlrd.open_workbook(location)
sheet = wb2.sheet_by_index(0)
sheet1.write(0, 1, sheet.cell_value(1, 0))
sheet1.write(1, 1, sheet.cell_value(1, 1))
sheet1.write(0, 2, sheet.cell_value(19, 0))
sheet1.write(1, 2, sheet.cell_value(19, 13))
sheet1.write(0, 3, 'Narration')
sheet2 = wb2.sheet_by_index(6)
i = int(2)
j = int(1)
k = int(1)
l = int(3)
while j<15:
	while i<11:
		sheet1.write(k, l, sheet2.cell_value(i, j))
		k+1
		i+1
	l+1
        
