from ColumnConverter import num_hash
from path import SheetNames, File

def find_card(codeORname):
    if (codeORname == "N"):
        Card_Name = input("Enter the name of the card you wish to find: ").upper() # Search result is Case sensitive 

        for sheet in SheetNames:
            currentSheet = DisplayCurrentSheet(sheet)
            find_cardName(currentSheet, Card_Name)

    if (codeORname == "C"):
        Card_Code = input("Enter the code of the card you wish to find: ").upper() # Search result is Case sensitive

        for sheet in SheetNames:
            currentSheet = DisplayCurrentSheet(sheet)
            find_cardCode(currentSheet, Card_Code)

def DisplayCurrentSheet(sheet):
    print("\nCurrent sheet is: {}".format(sheet) + "\n")
    currentSheet = File[sheet]
    return currentSheet

def find_cardName(currentSheet, Card_Name):
    for row in range(1, currentSheet.max_row + 1):
        for column in range(1,currentSheet.max_column + 1): #columns
            cell1Column = num_hash(column)
            cell2Column = num_hash(column - 1)
            cell3Column = num_hash(column + 1)
            cell_name = "{}{}".format(cell1Column, row)
            card_code_cell = "{}{}".format(cell2Column, row)
            card_quantity_cell = "{}{}".format(cell3Column, row)
            if currentSheet[cell_name].value == Card_Name: #Can't manipulate the cell value so can't use upper(), title() etc...
                print("There are {} {} {}. It is at cell {}".format(currentSheet[card_quantity_cell].value, currentSheet[cell_name].value, currentSheet[card_code_cell].value ,cell_name))

def find_cardCode(currentSheet, Card_Code):
    for row in range(1, currentSheet.max_row + 1):
        for column in range(1, currentSheet.max_column + 1): #columns
            cell1Column = num_hash(column)
            cell2Column = num_hash(column + 1)
            cell3Column = num_hash(column + 2)
            cell_name = "{}{}".format(cell1Column, row)
            card_code_cell = "{}{}".format(cell2Column, row)
            card_quantity_cell = "{}{}".format(cell3Column, row)
            if currentSheet[cell_name].value == Card_Code: #Can't manipulate the cell value so can't use upper(), title() etc...
                print("There are {} {} {}. It is at cell {}".format(currentSheet[card_quantity_cell].value, currentSheet[cell_name].value, currentSheet[card_code_cell].value ,cell_name))