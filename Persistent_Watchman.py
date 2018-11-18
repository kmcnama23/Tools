import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

#number=input('Enter a number: ')
formula_sheet = xlsxwriter.Workbook('C:/Users/Kyle McNamara/Desktop/Work/Gatekeeper_Information/Staging Area/formula_Sheet.xlsx')
ws1=formula_sheet.add_worksheet("Gatekeeper_Forumlas")
l=1
#for row_num in range(1,row_count):
#	hist_key = []
#	brand = '$G%d' % (row_num)
#	code = '$N%d' % (row_num)
#	endate = '$Q%d' % (row_num)
#	concat_values = [[str(brand),str(code),str(endate)]] 
#	for row in range(2, RSC137.max_row + 1):
#		hist_key.append({key:RSC137[col+str(row_num)].value for brand, code, endate in concat_values})
#	print(hist_key)
	
#for row_num in range(1,1000):
#	hist_key = xl_rowcol_to_cell(row_num,1)
#	concat = '=$G%d +|+ $N%d +|+ $Q%d)' % (row_num,row_num,row_num)
#	ws1.write(row_num,0,concat)

for row_num in range(1,1000):
	RDFB = xl_rowcol_to_cell(row_num,1)
	nearest_five = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row_num,row_num,row_num)
	ws1.write(row_num,0,nearest_five)
	
for row_num in range(1,1000):
	lift = xl_rowcol_to_cell(row_num,1)
	one = int(l)
	ws1.write(row_num,1,one)
	
for row_num in range(1,1000):
	billed = xl_rowcol_to_cell(row_num,1)
	shipped = '=IFERROR(VLOOKUP($B%d,GM History!A:AG,33,FALSE),"")' % (row_num)
	ws1.write(row_num,2,shipped)
	
for row_num in range(1,1000):
	cuts = xl_rowcol_to_cell(row_num,1)
	oos = '=IFERROR(VLOOKUP($B%d,GM History!A:AK,37,FALSE),"")' % (row_num)
	ws1.write(row_num,3,oos)
	
for row_num in range(1,1000):
	distros = xl_rowcol_to_cell(row_num,1)
	forced_out = '=IFERROR(VLOOKUP($B%d,GM History!A:AJ,36,FALSE),"")' % (row_num)
	ws1.write(row_num,4,forced_out)

for row_num in range(1,1000):
	sales = xl_rowcol_to_cell(row_num,1)
	pos = '=VLOOKUP($C%d,GM Sales!A:K,11,FALSE)' % (row_num)
	ws1.write(row_num,5,pos)

for row_num in range(1,1000):
	billed_sales = xl_rowcol_to_cell(row_num,1)
	ratio = '=$T%d/$W%d' % (row_num,row_num)
	ws1.write(row_num,6,ratio)

for row_num in range(1,1000):
	revised_Booking = xl_rowcol_to_cell(row_num,1)
	revision = '=IF($R%d>$P%d,$R%d,"")' % (row_num,row_num,row_num)
	ws1.write(row_num,7,revision)	
	
for row_num in range(1,1000):
	gk_booking = xl_rowcol_to_cell(row_num,1)
	record = '=IF($R%d>$P%d,$R%d,$P%d)' % (row_num,row_num,row_num,row_num)
	ws1.write(row_num,8,record)

for row_num in range(1,1000):
	max_billed = xl_rowcol_to_cell(row_num,1)
	moved = '=VLOOKUP($B%d,GM History!A:AP,42,FALSE)' % (row_num)
	ws1.write(row_num,9,moved)
	
for row_num in range(1,1000):
	max_sales = xl_rowcol_to_cell(row_num,1)
	sold = '=VLOOKUP($C%d,GM Sales!A:L,12,FALSE)' % (row_num)
	ws1.write(row_num,10,sold)
	
formula_sheet.close()

