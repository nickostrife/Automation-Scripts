import openpyxl
import os
from pathlib import Path

wb = openpyxl.load_workbook(Path('C:/Users/Nicko/Documents/automate_online-materials/example.xlsx'))
print(wb.sheetnames)
sheet = wb['Sheet3']
print(sheet)
print(type(sheet))
print(sheet.title)
anotherSheet = wb.active
print(anotherSheet)