import xlsxwriter

ad_review_week = input("Which ad week are you reviewing? ")
Gatekeeper = xlsxwriter.Workbook("C:/Users/Kyle McNamara/Desktop/Work/Code/Gatekeeper/Gatekeeper_Wk"+ad_review_week+"_Working_File.xlsx") 
ws1 = Gatekeeper.add_worksheet()	

Gatekeeper.close()	

