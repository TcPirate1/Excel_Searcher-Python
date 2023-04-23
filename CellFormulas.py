from ColumnConverter import num_hash
from path import File, Path
import re
from openpyxl.styles import Font, Alignment, NamedStyle

def find_cardLocation(currentSheet, Card):
    cardNameRegex = re.match('^\d{1,2}-\d{3}[CRHLS]+$', Card)
    for row in range(1, currentSheet.max_row + 1):
        for column in range(1,currentSheet.max_column + 1): #columns
            initialColumn = num_hash(column)
            leftColumn = num_hash(column - 1)
            rightColumn = num_hash(column + 1)
            secondRightColumn = num_hash(column + 2)
            searchTarget = f"{initialColumn}{row}"
            pile = f"{initialColumn}2"
            left_cell = f"{leftColumn}{row}"
            right_cell = f"{rightColumn}{row}"
            right_2Cells = f"{secondRightColumn}{row}"

            if (currentSheet[searchTarget].value == Card and cardNameRegex is None): #Can't manipulate the cell value so can't use upper(), title() etc...
                print(f"There are {currentSheet[right_cell].value} {currentSheet[left_cell].value} {currentSheet[searchTarget].value}(s). It is in the {currentSheet[pile].value} pile at Cell {searchTarget}.")

            if (currentSheet[searchTarget].value == Card and cardNameRegex is not None): #re.match returns None if no matches are found
                print(f"There are {currentSheet[right_2Cells].value} {currentSheet[searchTarget].value} {currentSheet[right_cell].value}(s) at Cell {searchTarget}.")

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
    # Create NamedStyle for reusability
    cell_format = NamedStyle(name = 'apply_format')
    cell_format.font = Font(bold=True)
    cell_format.alignment = Alignment(horizontal='center')
    # Input location of cell and get the 2 cells next to them
    selected_cell = input("Choose the cell you want to write to:\n").upper()
    selected_cellRight1 = ActiveWorksheet[selected_cell].offset(row= 0, column = 1).coordinate
    selected_cellRight2 = ActiveWorksheet[selected_cell].offset(row= 0, column = 2).coordinate
    # Input value for cell
    change_value = input(f"Enter what you want to put in cell {selected_cell}:\n").upper()
    change_valueRight1 = input(f"Enter what you want to put in this cell {selected_cellRight1}:\n").upper()
    change_valueRight2 = input(f"Enter what you want to put in this cell {selected_cellRight2}:\n").upper()
    #Assign new value
    ActiveWorksheet[selected_cell].value = change_value
    ActiveWorksheet[selected_cellRight1].value = change_valueRight1
    ActiveWorksheet[selected_cellRight2].value = change_valueRight2
    #Add checker so that it adds style to workbook if it isn't avaliable (like for instance on another PC)
    if 'apply_format' not in File:
        File.add_named_style(cell_format)
    #Apply style to cell after knowing where the cell is
    ActiveWorksheet[selected_cell].style = 'apply_format'
    ActiveWorksheet[selected_cellRight1].style = 'apply_format'
    ActiveWorksheet[selected_cellRight2].style = 'apply_format'
    # Sucessful output
    print(f"\"{ActiveWorksheet[selected_cell].value}\", \"{ActiveWorksheet[selected_cellRight1].value}\", \"{ActiveWorksheet[selected_cellRight2].value}\" has successfully been inputted into the spreadsheet.")
    File.save(Path)

def reapply_conditional_formatting(): ##Reapply conditional formatting rules after deleting values
    pass