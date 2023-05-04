from ColumnConverter import num_hash
import re
import time

def getInput(msg):
    return input(msg).upper()

def find_cardLocation(currentSheet, Card, searchType):
    s_time = time.time()
    cardNameRegex = re.match('^\d{1,2}-\d{3}[CRHLS]+$', Card)
    for row in range(1, currentSheet.max_row + 1):
        for column in range(1,currentSheet.max_column + 1): #columns
            startColumn = num_hash(column)
            leftColumn = num_hash(column - 1)
            rightColumn = num_hash(column + 1)
            secondRightColumn = num_hash(column + 2)
            searchTarget = f"{startColumn}{row}"
            pile = f"{startColumn}2"
            left_cell = f"{leftColumn}{row}"
            right_cell = f"{rightColumn}{row}"
            right_2Cells = f"{secondRightColumn}{row}"

            if (currentSheet[searchTarget].value == Card and cardNameRegex is None and searchType == "name"): #Can't manipulate the cell value so can't use upper(), title() etc...
                print(f"There are {currentSheet[right_cell].value} {currentSheet[left_cell].value} {currentSheet[searchTarget].value}(s). It is in the {currentSheet[pile].value} pile at Cell {searchTarget}.")
                print(f"\nTime taken: {time.time() - s_time} seconds.\n")

            if (currentSheet[searchTarget].value == Card and cardNameRegex is not None and searchType == "code"): #re.match returns None if no matches are found
                print(f"There are {currentSheet[right_2Cells].value} {currentSheet[searchTarget].value} {currentSheet[right_cell].value}(s) at Cell {searchTarget}.")
                print(f"\nTime taken: {time.time() - s_time} seconds.\n")