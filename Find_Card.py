from ColumnConverter import num_hash
from path import SheetNames, File
import re

def find_card(codeORname):
    if (codeORname == "N"):
        Card_Name = input("Enter the name of the card you wish to find: ").upper() # Search result is Case sensitive 

        for sheet in SheetNames:
            currentSheet = DisplayCurrentSheet(sheet)
            find_cardLocation(currentSheet, Card_Name)

    if (codeORname == "C"):
        Card_Name = input("Enter the code of the card you wish to find: ").upper() # Search result is Case sensitive

        for sheet in SheetNames:
            currentSheet = DisplayCurrentSheet(sheet)
            find_cardLocation(currentSheet, Card_Name)

def DisplayCurrentSheet(sheet):
    print("\nCurrent sheet is: {}".format(sheet) + "\n")
    currentSheet = File[sheet]
    return currentSheet

def find_cardLocation(currentSheet, Card_Name):
    cardNameRegex = re.match('^\d{1,2}-\d{3}[CRHLS]+$', Card_Name)
    for row in range(1, currentSheet.max_row + 1):
        for column in range(1,currentSheet.max_column + 1): #columns
            cell1Column = num_hash(column)
            cell2Column = num_hash(column - 1)
            cell3Column = num_hash(column + 1)
            cell4Column = num_hash(column + 2)
            searchTarget = "{}{}".format(cell1Column, row)
            left_cell = "{}{}".format(cell2Column, row)
            right_cell = "{}{}".format(cell3Column, row)
            right_2Cells = "{}{}".format(cell4Column, row)

            if (currentSheet[searchTarget].value == Card_Name and cardNameRegex == None): #Can't manipulate the cell value so can't use upper(), title() etc...
                print(f"There are {currentSheet[right_cell].value} {currentSheet[searchTarget].value} {currentSheet[left_cell].value}. It is at cell {searchTarget}")

            if (currentSheet[searchTarget].value == Card_Name and cardNameRegex != None): #re.match returns None if no matches are found
                print(f"There are {currentSheet[right_2Cells].value} {currentSheet[searchTarget].value} {currentSheet[right_cell].value}. It is at cell {searchTarget}")