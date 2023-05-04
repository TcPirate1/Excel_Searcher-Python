from path import SheetNames, File
from CellFormulas import find_cardLocation, getInput

def find_card(codeORname):
    Card = ""
    searchType = ""
    #initialize with empty strings then assign respective values to eliminate extra for loop

    if (codeORname == "N"):
        Card = getInput("Enter the name of the card you wish to find: ") # Search result is Case sensitive
        searchType = "name"

    if (codeORname == "C"):
        Card = getInput("Enter the code of the card you wish to find: ") # Search result is Case sensitive
        searchType = "code"

    for sheet in SheetNames:
        currentSheet = DisplayCurrentSheet(sheet)
        find_cardLocation(currentSheet, Card, searchType)

def DisplayCurrentSheet(sheet):
    print(f"\nCurrent sheet is: {sheet}\n")
    currentSheet = File[sheet]
    return currentSheet