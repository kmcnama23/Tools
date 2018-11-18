import win32com.client as win32

ad_review_week = input("Which ad week are you reviewing?(one last time) ")	
destination_path='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Code\\Gatekeeper\\Gatekeeper_Wk'+ad_review_week+'_Working_File.xlsx'

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible=False;

wb = excel.Workbooks.Open(destination_path);
ws = wb.Worksheets("GM Wk"+ad_review_week)

lastRow=ws.UsedRange.Rows.Count

col = ws.Range("B2:B500")
for cell in col:
	cell.Offset(1,1).Value = None



	
	
wb.Close(SaveChanges=True)
