import openpyxl
import os 
import os.path 
from openpyxl import Workbook
# get path  
file_path = os.getcwd()
# Give the location of the file
"""
library https://openpyxl.readthedocs.io/en/latest/usage.html
"""
# workbook object is created read file excel 
wb_obj = openpyxl.load_workbook(file_path + "/data.xlsx")
sheet_obj = wb_obj.active
m_row = sheet_obj.max_row
# Loop will print all values
# of first column
for i in range(1, m_row + 1):
	cell_obj = sheet_obj.cell(row = i, column = 1)
	print(cell_obj.value)
# create sheet 
worksheet = Workbook()
sheet1 = worksheet.create_sheet("Sheet1")
# write file excel 
