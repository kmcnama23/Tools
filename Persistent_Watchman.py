import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

formula_sheet = xlsxwriter.Workbook("C:/Users/Kyle McNamara/Desktop/Work/Code/Gatekeeper/formula_Sheet.xlsx")
ws1=formula_sheet.add_worksheet("Sheet1")

for row_num in range(1,10):
	RDFB = xl_rowcol_to_cell(row_num,1)
	nearest_five = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row_num,row_num,row_num)
	ws1.write(row_num,0,nearest_five)
	
for row_num in range(2,10):
	RDFB = xl_rowcol_to_cell(row_num,1)
	nearest_five = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row_num,row_num,row_num)
	ws1.write(row_num,0,nearest_five)
	
for row_num in range(3,10):
	RDFB = xl_rowcol_to_cell(row_num,1)
	nearest_five = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row_num,row_num,row_num)
	ws1.write(row_num,0,nearest_five)

	
formula_sheet.close()

