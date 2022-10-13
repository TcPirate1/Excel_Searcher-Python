from ColumnConverter import num_hash
from path import SheetNames, File

def find_cardBYname():
    Card_Name = input("Enter the name of the card you wish to find: ").title() # Search result is Case sensitive

    for sheet in SheetNames:
        print("\nCurrent sheet is: {}".format(sheet) + "\n")
        currentSheet = File[sheet]
        find_cardName(currentSheet, Card_Name)

def find_cardBYcode():
    Card_Code = input("Enter the name of the card you wish to find: ").title() # Search result is Case sensitive

    for sheet in SheetNames:
        print("\nCurrent sheet is: {}".format(sheet) + "\n")
        currentSheet = File[sheet]
        find_cardCode(currentSheet, Card_Code)

def find_cardName(currentSheet, Card_Name):
    for row in range(1, currentSheet.max_row + 1):
        for column in range(1,703): #columns
            cell1Column = num_hash(column)
            cell2Column = num_hash(column - 1)
            cell3Column = num_hash(column + 1)
            cell_name = "{}{}".format(cell1Column, row)
            cell_name2 = "{}{}".format(cell2Column, row)
            cell_name3 = "{}{}".format(cell3Column, row)
            if currentSheet[cell_name].value == Card_Name: #Can't manipulate the cell value so can't use upper(), title() etc...
                print("There are {} {} {}. It is at cell {}".format(currentSheet[cell_name3].value, currentSheet[cell_name].value, currentSheet[cell_name2].value ,cell_name))

def find_cardCode(currentSheet, Card_Code):
    for row in range(1, currentSheet.max_row + 1):
        for column in range(1,703): #columns
            cell1Column = num_hash(column)
            cell2Column = num_hash(column + 1)
            cell3Column = num_hash(column + 2)
            cell_name = "{}{}".format(cell1Column, row)
            cell_name2 = "{}{}".format(cell2Column, row)
            cell_name3 = "{}{}".format(cell3Column, row)
            if currentSheet[cell_name].value == Card_Code: #Can't manipulate the cell value so can't use upper(), title() etc...
                print("There are {} {} {}. It is at cell {}".format(currentSheet[cell_name3].value, currentSheet[cell_name].value, currentSheet[cell_name2].value ,cell_name))