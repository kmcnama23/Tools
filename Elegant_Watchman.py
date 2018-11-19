import win32com.client as win32

ad_review_week = input("Which ad week are you reviewing?(one last time) ")	
destination_path='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Code\\Gatekeeper\\Gatekeeper_Wk'+ad_review_week+'_Working_File.xlsx'

xl = win32.gencache.EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open('C:\\Users\\Kyle McNamara\\Desktop\\Work\\Code\\Gatekeeper\\Gatekeeper_Wk'+ad_review_week+'_Working_File.xlsx')
ws = wb.Worksheets('GM Wk'+ad_review_week)
xl.Visible=True
wb.Close(savechanges=1)
xl.Quit()




