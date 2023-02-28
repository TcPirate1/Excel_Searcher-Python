from ColumnConverter import num_hash
import re

def find_cardLocation(currentSheet, Card):
    cardNameRegex = re.match('^\d{1,2}-\d{3}[CRHLS]+$', Card)
    for row in range(1, currentSheet.max_row + 1):
        for column in range(1,currentSheet.max_column + 1): #columns
            cell1Column = num_hash(column)
            cell2Column = num_hash(column - 1)
            cell3Column = num_hash(column + 1)
            cell4Column = num_hash(column + 2)
            searchTarget = f"{cell1Column}{row}"
            pile = f"{cell1Column}2"
            left_cell = f"{cell2Column}{row}"
            right_cell = f"{cell3Column}{row}"
            right_2Cells = f"{cell4Column}{row}"

            if (currentSheet[searchTarget].value == Card and cardNameRegex == None): #Can't manipulate the cell value so can't use upper(), title() etc...
                print(f"There are {currentSheet[right_cell].value} {currentSheet[searchTarget].value} {currentSheet[left_cell].value}. It is in the {currentSheet[pile].value} pile")

            if (currentSheet[searchTarget].value == Card and cardNameRegex != None): #re.match returns None if no matches are found
                print(f"There are {currentSheet[right_2Cells].value} {currentSheet[searchTarget].value} {currentSheet[right_cell].value}. It is in the {currentSheet[pile].value} pile")