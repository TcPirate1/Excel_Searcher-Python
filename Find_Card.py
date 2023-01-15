from path import SheetNames, File
from CellFormulas import find_cardLocation

def find_card(codeORname):
    if (codeORname == "N"):
        Card = input("Enter the name of the card you wish to find: ").upper() # Search result is Case sensitive 

        for sheet in SheetNames:
            currentSheet = DisplayCurrentSheet(sheet)
            find_cardLocation(currentSheet, Card)

    if (codeORname == "C"):
        Card = input("Enter the code of the card you wish to find: ").upper() # Search result is Case sensitive

        for sheet in SheetNames:
            currentSheet = DisplayCurrentSheet(sheet)
            find_cardLocation(currentSheet, Card)

def DisplayCurrentSheet(sheet):
    print(f"\nCurrent sheet is: {sheet}\n")
    currentSheet = File[sheet]
    return currentSheet