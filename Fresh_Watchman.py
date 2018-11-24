import xlsxwriter
import openpyxl
import openpyxl as xl
from openpyxl import Workbook
from openpyxl import styles
from openpyxl.styles import Font, PatternFill, Alignment, Color
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import coordinate_from_string, column_index_from_string
from shutil import copyfile
from abc import ABCMeta, abstractmethod


ad_review_week = input("Which ad week are you reviewing? ")
Gatekeeper = xlsxwriter.Workbook("C:/Users/Kyle McNamara/Desktop/Work/Code/Gatekeeper/Gatekeeper_Wk"+ad_review_week+"_Working_File.xlsx") 
gk = Gatekeeper.add_worksheet("GM Wk"+ad_review_week)
hist = Gatekeeper.add_worksheet("GM History")
sales = Gatekeeper.add_worksheet("GM Sales")
Gatekeeper.close()
source_path = 'C:\\Users\\Kyle McNamara\\Desktop\\Work\\Gatekeeper_Information\\Week '+ad_review_week+', 2019\\GM_Wk'+ad_review_week+'_GK_File.xlsx'
destination_path = 'C:\\Users\\Kyle McNamara\\Desktop\\Work\\Code\\Gatekeeper\\Gatekeeper_Wk'+ad_review_week+'_Working_File.xlsx'
hist_sales_path ='C:\\Users\\Kyle McNamara\\Desktop\\Work\\Gatekeeper_Information\\Week '+ad_review_week+', 2019\\GM_History_&_Sales.xlsx'	
gkdata = xl.load_workbook(filename=source_path)
wrkngfile=xl.load_workbook(filename=destination_path)
histdata = xl.load_workbook(filename=hist_sales_path)
sourcedata = gkdata["GM Wk"+ad_review_week]
billdata=histdata["GM History"]
salesdata=histdata["GM Sales"]
gktarget=wrkngfile["GM Wk"+ad_review_week]
histtarget=wrkngfile["GM History"]
salestarget=wrkngfile["GM History"] 
sheet=gkdata["GM Wk"+ad_review_week]
active_rows = gkdata.worksheets[0]
row_count = active_rows.max_row
col_count = active_rows.max_column

class DataImport(object):
	
	
	def __init__(self, startCol, startRow, endCol, endRow, sheet):
		self.startCol=startCol
		self.startRow=startRow
		self.endCol=endCol
		self.endRow=endRow
		
	def gkCopy(self):
		startCol=1
		startRow=1
		endCol=col_count
		endRow=row_count
		rangeSelected = []
		#Loops through selected Rows
		for i in range(startRow,endRow + 1,1):
		#Appends the row to a RowSelected list
			rowSelected = []
			for j in range(startCol,endCol+1,1):
				rowSelected.append(sheet.cell(row = i, column = j).value)
			#Adds the RowSelected List and nests inside the rangeSelected
			rangeSelected.append(rowSelected)
		return rangeSelected

ingest = DataImport(1,1,col_count,row_count,sourcedata)
print(gkCopy.ingest())		
		
class DataSet(object):

	def __init__(ingest,digest,absorb):
		self.ingest = ingest
		self.digest = digest
		self.absorb = absorb
	
	def gkTab(ingest):
		self.ingest = DataImport(1,1,col_count,row_count,sourcedata)
		
	def histTab(digest):
		self.digest = DataImport(1,1,col_count,row_count,billdata)
		
	def salesTab(absorb):
		self.absorb = DataImport(1,1,col_count,row_count,salesdata)
		
	def grab(self):
		return self.ingest
	
	def hold(self):
		return self.digest
		
	def squeeze(self):
		return self.absorb
		

print(gold)


		


	
#def paste(firstCol, firstRow, lastCol, lastRow, sheetReceiving,copiedData):
#	numberRow = 0
#	for m in range(firstRow,lastRow+1,1):
#		numberCol = 0
#		for o in range(firstCol,lastCol+1,1):
#		
#			sheetReceiving.cell(row = m , column = o).value = copiedData[numberRow][numberCol]
#			numberCol += 1
#		numberRow += 1
		
#def input():
#	gkContent = ingest.gkCopy()
#	histContent = digest.gkCopy()
#	salesContent = absorb.gkCopy()
#	gkDestination=paste(1,1,col_count,row_count,gktarget,gkContent)
#	histDestination=paste(1,1,col_count,row_count,histtarget,histContent)
#	salesDestination=paste(1,1,col_count,row_count,salestarget,salesContent)
#input()		

wrkngfile.save(destination_path)

