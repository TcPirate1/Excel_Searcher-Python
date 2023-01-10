import openpyxl
import os

Path = os.path.join(os.path.expanduser('~'), 'Documents', 'Spreadsheets', 'FFTCG_automated.xlsx')
File = openpyxl.load_workbook(Path, data_only=False)
SheetNames = File.sheetnames