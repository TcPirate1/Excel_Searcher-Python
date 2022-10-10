import openpyxl
import os
from path import RootPath

def find_specific_cell():
    for row in range(1, currentSheet.max_row + 1):
        for column in range(1,703): #Rows
            column = num_hash(column)
            cell_name = "{}{}".format(column, row)
            if currentSheet[cell_name.capitalize()].value == Card_Name: #Case sensitive
                print("{} is at {}".format(currentSheet[cell_name].value, cell_name))

def num_hash(num):
    if num < 26:
        return alpha[num-1]
    else:
        quotient, remainder = num // 26, num % 26
        if remainder == 0:
            if quotient == 1:
                return alpha[remainder-1]
            else:
                return num_hash(quotient-1) + alpha[remainder-1]
        else:
            return num_hash(quotient) + alpha[remainder-1]

#Variables
relativePath = os.path.join(RootPath, "FFTCG.xlsx")
File = openpyxl.load_workbook(relativePath, data_only=True)
SheetNames = File.sheetnames
Card_Name = input("Enter the code of the card you wish to find: ").capitalize()
alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

#Entry point
for sheet in SheetNames:
    print("\nCurrent sheet is: {}".format(sheet) + "\n")
    currentSheet = File[sheet]
    find_specific_cell()