import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

formula_sheet = xlsxwriter.Workbook('C:/Users/Kyle McNamara/Desktop/Work/Gatekeeper_Information/Staging Area/formula_Sheet.xlsx')
ws1=formula_sheet.add_worksheet("Sheet1")
l = 1

#for row_num in range(1,1000):
#	hist_key = xl_rowcol_to_cell(row_num,1)
#	concat = '=concat($G%d +|+ $N%d +|+ $Q%d)' % (row_num,row_num,row_num)
#	ws1.write(row_num,0,concat)

for row_num in range(1,1000):
	RDFB = xl_rowcol_to_cell(row_num,1)
	nearest_five = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row_num,row_num,row_num)
	ws1.write(row_num,1,nearest_five)
	
for row_num in range(1,1000):
	lift = xl_rowcol_to_cell(row_num,1)
	one = int(l)
	ws1.write(row_num,2,one)
	
#for row_num in range(4,1000):
#	billed = xl_rowcol_to_cell(row_num,1)
#	nearest_five = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row_num,row_num,row_num)
#	ws1.write(row_num,3,nearest_five)
	
#for row_num in range(5,1000):
#	cuts = xl_rowcol_to_cell(row_num,1)
#	nearest_five = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row_num,row_num,row_num)
#	ws1.write(row_num,4,nearest_five)
	
#for row_num in range(6,1000):
#	distros = xl_rowcol_to_cell(row_num,1)
#	nearest_five = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row_num,row_num,row_num)
#	ws1.write(row_num,5,nearest_five)

#for row_num in range(7,1000):
#	sales = xl_rowcol_to_cell(row_num,1)
#	pos = '=$T%d/$W%d' % (row_num,row_num)
#	ws1.write(row_num,6,ratio)

for row_num in range(1,1000):
	billed_sales = xl_rowcol_to_cell(row_num,1)
	ratio = '=$T%d/$W%d' % (row_num,row_num)
	ws1.write(row_num,7,ratio)

for row_num in range(1,1000):
	revised_Booking = xl_rowcol_to_cell(row_num,1)
	revision = '=IF($R%d>$P%d,$R%d,"")' % (row_num,row_num,row_num)
	ws1.write(row_num,8,revision)	
	
for row_num in range(1,1000):
	gk_booking = xl_rowcol_to_cell(row_num,1)
	record = '=IF($R%d>$P%d,$R%d,$P%d)' % (row_num,row_num,row_num,row_num)
	ws1.write(row_num,9,record)

for row_num in range(1,1000):
	max_billed = xl_rowcol_to_cell(row_num,1)
	moved = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row_num,row_num,row_num)
	ws1.write(row_num,10,moved)
	
for row_num in range(1,1000):
	max_sales = xl_rowcol_to_cell(row_num,1)
	sold = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row_num,row_num,row_num)
	ws1.write(row_num,11,sold)
	
formula_sheet.close()

