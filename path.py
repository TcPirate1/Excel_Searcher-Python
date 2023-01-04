import openpyxl
import os

Path = os.path.join(os.environ['USERPROFILE'], 'Documents', 'Spreadsheets', 'FFTCG.xlsx')
File = openpyxl.load_workbook(Path, data_only=True)
SheetNames = File.sheetnames