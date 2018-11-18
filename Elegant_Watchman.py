import openpyxl

ad_review_week = input("Which ad week are you reviewing?(one last time) ")	
destination_path='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Code\\Gatekeeper\\Gatekeeper_Wk'+ad_review_week+'_Working_File.xlsx'

wb = openpyxl.load_workbook(destination_path)
ws1 = wb['GM Wk'+ad_review_week]
row_count = ws1.max_row

row=2
while row < row_count:
		for row,formula in enumerate(list(ws1.columns)[1],1):
			concat= '=G%d&"|"&N%d&"|"&Q%d' % (row,row,row)
			formula.value = concat
			row += 1
			

ws1['B1']="History Key"

		
wb.save(destination_path)
	

