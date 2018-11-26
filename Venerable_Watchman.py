from win32com.client import Dispatch

print('Beginning billed and sales history data import. Please Standby')
ad_review_week = input("Which ad week are you reviewing?(one last time) ")	
destination_path='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Code\\Gatekeeper\\Gatekeeper_Wk'+ad_review_week+'_Working_File.xlsx'
hist_sales_path ='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Gatekeeper_Information\\Week '+ad_review_week+', 2019\\GM_History_&_Sales.xlsx'

xl = Dispatch("Excel.Application")
xl.Visible=False

wb1 = xl.Workbooks.Open(Filename=hist_sales_path)
wb2 = xl.Workbooks.Open(Filename=destination_path)

ws1=wb1.Worksheets(1)
ws2=wb1.Worksheets(2)
ws1.Copy(wb2.Worksheets(1))
ws2.Copy(wb2.Worksheets(2))


wb = xl.Workbooks.Open(Filename=destination_path, ReadOnly='False')
for worksheet in wb.Sheets:
	if worksheet.Name == 'GM GateKeeper':
		worksheet.Move(Before=wb.Sheets("GM History"))
print("Data import and format complete. Please use Elegant_Watchman.py to complete Gatekeeper build.") 
		
		
wb2.Close(SaveChanges=True)
xl.Quit()


