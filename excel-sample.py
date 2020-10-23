import openpyxl
import os
from pathlib import Path

wb = openpyxl.load_workbook(Path('C:/Users/Nicko/Documents/automate_online-materials/example.xlsx'))
print(type(wb))