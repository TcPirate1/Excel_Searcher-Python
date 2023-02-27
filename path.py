import openpyxl
import os

Path = os.path.join(os.path.expanduser('~'), 'Documents', 'Spreadsheets', 'FFTCG')
File = openpyxl.load_workbook(Path, data_only=True)
SheetNames = File.sheetnames