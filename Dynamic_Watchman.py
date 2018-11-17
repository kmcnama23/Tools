import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

ad_review_week = input("Which ad week are you reviewing? ")
fa_name=input("Please enter your name: ")
print("Spinning Gatekeeper Week,"+ad_review_week+" Instance for "+fa_name)
Gatekeeper = xlsxwriter.Workbook("C:/Users/Kyle McNamara/Desktop/Work/Code/Gatekeeper/Gatekeeper_Wk"+ad_review_week+"_Working_File.xlsx") 
ws1 = Gatekeeper.add_worksheet()
print("File creation successful. Please initialize Venerable_Watchman.py on the next line to complete Gatekeeper build.")	

Gatekeeper.close()	

