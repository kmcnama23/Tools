import win32com.client
from win32com.client import Dispatch
from win32com.client import constants
import shutil

ad_review_week = input("Which ad week are you reviewing? ")
fa_name=input("Please enter your name: ")
shutil.copy('C:\\Users\\Kyle McNamara\\Desktop\\Work\\Gatekeeper_Information\\Week '+ad_review_week+', 2019\\GM_Wk'+ad_review_week+'_GK_File.xlsx',
			'C:\\Users\\Kyle McNamara\\Desktop\\Work\\Code\\Gatekeeper\\FileBuild.xlsx')
			
print('Beginning billed and sales history data import. Please Standby')	
destination_path='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Code\\Gatekeeper\\FileBuild.xlsx'
hist_sales_path ='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Gatekeeper_Information\\Week '+ad_review_week+', 2019\\GM_History_&_Sales.xlsx'


xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
xl.Visible=False
xl.ScreenUpdating=False
wb1 = xl.Workbooks.Open(Filename=hist_sales_path)
wb2 = xl.Workbooks.Open(Filename=destination_path,ReadOnly=0)

wb2_sheet=wb2.Worksheets(1)
wb2_sheet.Name='Sheet1'

ws1=wb1.Worksheets(1)
ws2=wb1.Worksheets(2)
ws1.Copy(wb2.Worksheets(1))
ws2.Copy(wb2.Worksheets(2))

wb1.Close(SaveChanges=False)

for worksheet in wb2.Sheets:
	if worksheet.Name == 'Sheet1':
		worksheet.Move(Before=wb2.Sheets("GM History"))
		
ws3_return=wb2.Worksheets(1)
ws3_return.Name="GM GateKeeper"

print("Data import and format complete. Please use Elegant_Watchman.py to complete Gatekeeper build.") 

analyst=fa_name
row = 2
while True:
	xl.Range("A%d:AH%d" % (row,row)).Select()
	data=xl.ActiveCell.FormulaR1C1
	xl.Range("D%d" % row).Select()
	condition = xl.ActiveCell.FormulaR1C1
	
	if data == '':
		break
	elif condition != str(analyst):
		xl.Rows("%d:%d" % (row,row)).Select()
		xl.Selection.Delete(Shift=constants.xlUp)
	else:
		row+=1	

ws3_return.Columns("Q:Z").Insert()

ws3_return.Cells(1,17).Value="LRS"
ws3_return.Cells(1,18).Value="RDFB"
ws3_return.Cells(1,19).Value="Lift"
ws3_return.Cells(1,20).Value="Billed"
ws3_return.Cells(1,21).Value="OOS"
ws3_return.Cells(1,22).Value="Distros"
ws3_return.Cells(1,23).Value="Sales"
ws3_return.Cells(1,24).Value="Billed/Sales"
ws3_return.Cells(1,25).Value="GK Revised Booking"
ws3_return.Cells(1,26).Value="GK Booking"
ws3_return.Cells(1,27).Value="Max Sales"
ws3_return.Cells(1,28).Value="Max Billed"
ws3_return.Cells(1,29).Value="Comments"

#wb=xl.Workbooks.Open(Filename=destination_path,ReadOnly=0)

salesCode='''
Sub Insert_Sales()

Application.ScreenUpdating = False

Dim EndRow As Long

EndRow = Sheets(3).Cells(Rows.Count, "A").End(xlUp).Row
Sheets(3).Range("J3:J" & EndRow).Formula = "=VLOOKUP(A3,'GM GateKeeper'!C:Q,15,FALSE)"
Sheets(3).Range("K3:K" & EndRow).Formula = "=HLOOKUP(J3,M:O,P3,FALSE)"

End Sub
'''

mod = wb2.VBProject.VBComponents.Add(1)
mod.CodeModule.AddFromString(salesCode)
xl.Run("Insert_Sales")

gkCode = '''

Sub Insert_Formula()

Dim LastRow As Long

Application.ScreenUpdating = False

LastRow = Sheets(1).Range("B" & Rows.Count).End(xlUp).Row

Sheets(1).Range("Q1").Interior.Color = RGB(255,127,80)
Sheets(1).Range("R1").Interior.Color = RGB(0,51,102)
Sheets(1).Range("S1").Interior.Color = RGB(63,63,64)
Sheets(1).Range("T1").Interior.Color = RGB(41,43,55)
Sheets(1).Range("U1").Interior.Color = RGB(255,215,0)
Sheets(1).Range("V1").Interior.Color = RGB(204,204,204)
Sheets(1).Range("W1").Interior.Color = RGB(102,0,102)
Sheets(1).Range("X1").Interior.Color = RGB(198,226,255)
Sheets(1).Range("Y1").Interior.Color = RGB(6,85,53)
Sheets(1).Range("Z1").Interior.Color = RGB(128,0,0)

Sheets(1).Range("B2:B" & LastRow).Formula = "=CONCAT(G2,""|"",N2,""|"",Q2)"
Sheets(1).Range("R2:R" & LastRow).Formula = "=ROUND(S2*(T2+U2*.6)/5,0)*5"
Sheets(1).Range("S2:S" & LastRow).Formula = 1
Sheets(1).Range("T2:T" & LastRow).Formula = "=IFERROR(VLOOKUP(B2,'GM HISTORY'!A:AG,33,FALSE),"""")"
Sheets(1).Range("U2:U" & LastRow).Formula = "=IFERROR(VLOOKUP(B2,'GM HISTORY'!A:AK,37,FALSE),"""")"
Sheets(1).Range("V2:V" & LastRow).Formula = "=IFERROR(VLOOKUP(B2,'GM HISTORY'!A:AJ,36,FALSE),"""")"
Sheets(1).Range("W2:W" & LastRow).Formula = "=VLOOKUP(C2,'GM SALES'!A:K,11,FALSE)"
Sheets(1).Range("X2:X" & LastRow).Formula = "=T2/W2"
Sheets(1).Range("Y2:Y" & LastRow).Formula = "=IF(R2>P2,R2,"""")"
Sheets(1).Range("Z2:Z" & LastRow).Formula = "=IF(R2>P2,R2,P2)"
Sheets(1).Range("AA2:AA" & LastRow).Formula = "=IFERROR(VLOOKUP(B2,'GM HISTORY'!A:AP,42,FALSE),"""")"
Sheets(1).Range("AB2:AB" & LastRow).Formula = "=VLOOKUP(C2,'GM SALES'!A:L,12,FALSE)"
Sheets(1).Range("Q2:Q" & LastRow).Interior.Color = RGB(255,229,204)
Sheets(1).Range("R2:R" & LastRow).Interior.Color = RGB(229,204,255)
Sheets(1).Range("S2:S" & LastRow).Interior.Color = RGB(63,63,64)
Sheets(1).Range("T2:T" & LastRow).Interior.Color = RGB(224,224,224)
Sheets(1).Range("U2:U" & LastRow).Interior.Color = RGB(255,255,204)
Sheets(1).Range("V2:V" & LastRow).Interior.Color = RGB(224,224,224)
Sheets(1).Range("W2:W" & LastRow).Interior.Color = RGB(215,215,230)
Sheets(1).Range("X2:X" & LastRow).Interior.Color = RGB(204,229,255)
Sheets(1).Range("Y2:Y" & LastRow).Interior.Color = RGB(204,255,204)
Sheets(1).Range("Z2:Z" & LastRow).Interior.Color = RGB(255,204,204)
Sheets(1).Range("AA2:AA" & LastRow).Interior.Color = RGB(192,192,192)
Sheets(1).Range("AB2:AB" & LastRow).Interior.Color = RGB(192,192,192)
Sheets(1).Range("AC2:AC" & LastRow).Interior.Color = RGB(192,192,192)
Sheets(1).Range("S2:S" & LastRow).Font.Color = RGB(255,255,255)
Sheets(1).Range("S2:S" & LastRow).Font.Bold = True
End Sub


'''

mod = wb2.VBProject.VBComponents.Add(1)
mod.CodeModule.AddFromString(gkCode)
xl.Run("Insert_Formula")

xl.ActiveSheet.Columns.AutoFit()

wb2.SaveAs(destination_path)
wb2.Close(SaveChanges=True)		

