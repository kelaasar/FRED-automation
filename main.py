from openpyxl.workbook import Workbook
from openpyxl import load_workbook

#Create a workbook object
wb = Workbook()

#load spreadsheet
wb = load_workbook('data/financialData.xlsm')

# create an active worksheet 
wb.active = wb['Shortname Summary']
ws = wb.active


print("Hello World!")


