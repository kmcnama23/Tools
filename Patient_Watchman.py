import openpyxl

ad_review_week = input("Which ad week are you reviewing?(one last time) ")	
destination_path='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Code\\Gatekeeper\\Gatekeeper_Wk'+ad_review_week+'_Working_File.xlsx'

wb = openpyxl.load_workbook(destination_path)
ws1 = wb['GM Wk'+ad_review_week]
row_count = ws1.max_row
l=1
row=2

while row < row_count:
	for row,formula in enumerate(list(ws1.columns)[1],1):
		concat= '=G%d&"|"&N%d&"|"&Q%d' % (row,row,row)
		formula.value = concat
		row += 1
		
	for row,formula in enumerate(list(ws1.columns)[17],1):
		RDFB = '=ROUND($S%d*($T%d+$U%d*.6)/5,0)*5' % (row,row,row)
		formula.value = RDFB
		row += 1
	
	for row,formula in enumerate(list(ws1.columns)[18],1):
		lift = l
		formula.value = lift
		row += 1
	
	for row,formula in enumerate(list(ws1.columns)[19],1):
		billed = '=IFERROR(VLOOKUP($B%d,GM History!A:AG,33,FALSE),"")' % (row)
		formula.value = billed
		row += 1
		
	for row,formula in enumerate(list(ws1.columns)[20],1):
		cuts = '=IFERROR(VLOOKUP($B%d,GM History!A:AK,37,FALSE),"")' % (row)
		formula.value = cuts
		row += 1
		
	for row,formula in enumerate(list(ws1.columns)[21],1):
		distros = '=IFERROR(VLOOKUP($B%d,GM History!A:AJ,36,FALSE),"")' % (row)
		formula.value = distros
		row += 1
		
	for row,formula in enumerate(list(ws1.columns)[22],1):
		sales = '=VLOOKUP($C%d,GM Sales!A:K,11,FALSE))' % (row)
		formula.value = sales
		row += 1
		
	for row,formula in enumerate(list(ws1.columns)[23],1):
		Billed_Sales = '=$T%d/$W%d' % (row,row)
		formula.value = Billed_Sales
		row += 1
		
	for row,formula in enumerate(list(ws1.columns)[24],1):
		revised_booking = '=IF($R%d>$P%d,$R%d,"")' % (row,row,row)
		formula.value = revised_booking
		row += 1
		
	for row,formula in enumerate(list(ws1.columns)[25],1):
		gk_booking = '=IF($R%d>$P%d,$R%d,$P%d)' % (row,row,row,row)
		formula.value = RDFB
		row += 1
		
	for row,formula in enumerate(list(ws1.columns)[26],1):
		max_billed = '=VLOOKUP($B%d,GM History!A:AP,42,FALSE)' % (row)
		formula.value = max_billed
		row += 1
		
	for row,formula in enumerate(list(ws1.columns)[27],1):
		max_sales = '=VLOOKUP($C%d,GM Sales!A:L,12,FALSE)' % (row)
		formula.value = max_sales
		row += 1
		
		
ws1['B1']="History Key"
ws1['R1']="RDFB"
ws1['S1']="Lift"
ws1['T1']="Billed"
ws1['U1']="Cuts"
ws1['V1']="Distros"
ws1['W1']="Sales"
ws1['X1']="Billed/Sales"
ws1['Y1']="Revised Booking"
ws1['Z1']="GK Booking"
ws1['AA1']="Max Billed"
ws1['AB1']="Max Sales"

wb.save(destination_path)