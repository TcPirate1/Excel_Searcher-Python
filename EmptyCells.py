from ColumnConverter import num_hash

def emptyCells(ActiveWorksheet):
    print("Checking for empty cells in this sheet...\n")
    for row in range(1, ActiveWorksheet.max_row + 1):
        for column in range(1, ActiveWorksheet.max_column + 1):
            columnLetter = num_hash(column)
            top = f"{columnLetter}1"
            searchTarget = f"{columnLetter}{row}"
            if (ActiveWorksheet[searchTarget].value is None and ActiveWorksheet[top].value == "Code"):
                print(searchTarget)