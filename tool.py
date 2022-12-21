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
wb = openpyxl.Workbook()
# workbook object is created read file excel 

wb_obj = openpyxl.load_workbook(file_path + "/data.xlsx")

sheet_obj = wb_obj.active

m_row = sheet_obj.max_row

# write file excel 

array1 = ["a.c", "a.c", "a.c", "b.c", "c.c", "c.c"]

array2 = ["func1", "func11", "func12", "funcb", "funcc", "funccc"]

# for i in range(1, m_row + 1):
		
# 		cell_obj = (sheet_obj.cell(row = i, column = 2))
for j in range(len(array1)):
    cell_obj = (sheet_obj.cell(row = j+1, column = 2))
    #print(cell_obj.value)
    if cell_obj.value == None:
        cell_obj.value = "File Name:" + array1[j]
        print(cell_obj.value)
wb_obj.save(file_path + "\data.xlsx")





