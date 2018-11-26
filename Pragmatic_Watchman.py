import win32com.client
from win32com.client import DispatchEx

ad_review_week = input("Which ad week are you reviewing?(one last time) ")	
destination_path='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Code\\Gatekeeper\\Gatekeeper_Wk'+ad_review_week+'_Working_File.xlsx'
hist_sales_path ='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Gatekeeper_Information\\Week '+ad_review_week+', 2019\\GM_History_&_Sales.xlsx'
macro_path='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Gatekeeper_Information\\Staging Area\\Dynamic_Insert.xlsm'

xl=DispatchEx("Excel.Application")
xl.Visible=False
wb=xl.Workbooks.Open(Filename=destination_path,ReadOnly=0)

salesCode='''
Sub Insert_Sales()

Application.ScreenUpdating = False

Dim EndRow As Long

EndRow = Sheets(3).Cells(Rows.Count, "A").End(xlUp).Row
Sheets(3).Range("J3:J" & EndRow).Formula = "=VLOOKUP(A3,'GM GateKeeper'!C:Q,15,FALSE)"
Sheets(3).Range("K3:K" & EndRow).Formula = "=HLOOKUP(J3,M:O,P3,FALSE)"

End Sub
'''

mod = wb.VBProject.VBComponents.Add(1)
mod.CodeModule.AddFromString(salesCode)
xl.Run("Insert_Sales")

gkCode = '''

Sub Insert_Formula()

Dim LastRow As Long

Application.ScreenUpdating = False

LastRow = Sheets(1).Range("B" & Rows.Count).End(xlUp).Row

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

End Sub


'''

mod = wb.VBProject.VBComponents.Add(1)
mod.CodeModule.AddFromString(gkCode)
xl.Run("Insert_Formula")

wb.SaveAs(destination_path)
xl.Quit() 