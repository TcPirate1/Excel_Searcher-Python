import openpyxl
import os

relativePath = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'FFTCG.xlsx')
File = openpyxl.load_workbook(relativePath, data_only=True)
SheetNames = File.sheetnames