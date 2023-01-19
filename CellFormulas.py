from ColumnConverter import num_hash
from path import File, Path
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
            left_cell = f"{cell2Column}{row}"
            right_cell = f"{cell3Column}{row}"
            right_2Cells = f"{cell4Column}{row}"

            if (currentSheet[searchTarget].value == Card and cardNameRegex == None): #Can't manipulate the cell value so can't use upper(), title() etc...
                print(f"There are {currentSheet[right_cell].value} {currentSheet[searchTarget].value} {currentSheet[left_cell].value}. It is at cell {searchTarget}")

            if (currentSheet[searchTarget].value == Card and cardNameRegex != None): #re.match returns None if no matches are found
                print(f"There are {currentSheet[right_2Cells].value} {currentSheet[searchTarget].value} {currentSheet[right_cell].value}. It is at cell {searchTarget}")

def emptyCells(ActiveWorksheet):
    print("Checking for empty cells in this sheet...\n")
    for row in range(1, ActiveWorksheet.max_row + 1):
        for column in range(1, ActiveWorksheet.max_column + 1):
            columnLetter = num_hash(column)
            top = f"{columnLetter}1"
            searchTarget = f"{columnLetter}{row}"
            if (ActiveWorksheet[searchTarget].value is None and ActiveWorksheet[top].value == "Code"):
                print(searchTarget)

def fillEmptyCell(ActiveWorksheet):
    print("Reminder, empty cells are under the \"Code\" column!")
    selected_cell = input("Choose the cell you want to write to:\n").upper()
    selected_cellRight1 = ActiveWorksheet[selected_cell].offset(row= 0, column = 1).coordinate
    selected_cellRight2 = ActiveWorksheet[selected_cell].offset(row= 0, column = 2).coordinate
    change_value = input(f"Enter what you want to put in cell {selected_cell}:\n").upper()
    change_valueRight1 = input(f"Enter what you want to put in this cell {selected_cellRight1}:\n").upper()
    change_valueRight2 = input(f"Enter what you want to put in this cell {selected_cellRight2}:\n").upper()
    ActiveWorksheet[selected_cell].value = change_value
    ActiveWorksheet[selected_cellRight1].value = change_valueRight1
    ActiveWorksheet[selected_cellRight2].value = change_valueRight2
    print(f"{ActiveWorksheet[selected_cell].value}, {ActiveWorksheet[selected_cellRight1].value}, {ActiveWorksheet[selected_cellRight2].value}")
    File.save(Path)