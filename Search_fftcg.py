from Find_Card import find_card
from CellFormulas import emptyCells, fillEmptyCell, getInput
from path import File, Path

def inputChoice():
    userContinue = True
    while userContinue == True:
        option = input("\nPress one of the following keys to execute their respective functions.\n\nOptions:\n\"E\" to exit the program.\n\"S\" to search for a card.\n\"A\" to add a new worksheet to the workbook.\n\"F\" to fill an empty cell on a worksheet.\n\n\"V\" will show empty cells in the \"Code\" column.\n\"C\" to change the value of a cell or delete a cell.\n\n").upper()
        match option:
            case "E":
                print("The program will now close.\nThank you for using this Python script!")
                break

            case "S":
                cardFinder()

            case "A":
                new_worksheet = input("What do you want to name this worksheet?\n\nNote: If you don't enter anything it will default to sheet + number.\n").capitalize()
                File.create_sheet(f"{new_worksheet}")
                File.save(Path)
                print("Worksheet created.")

            case "V":
                currentWorksheet = File.active
                NoCellValue = getInput(f"Current sheet is: {currentWorksheet}\n\nWould you like to see the empty cells in {currentWorksheet}? (Y = yes, N = no)\n")
                if (NoCellValue == "Y"):
                    emptyCells(currentWorksheet)
                else:
                    print("Returning to Menu Screen.")
            
            case "F":
                print(f"Current sheet is: {File.active}")
                ChangeActiveWorksheet = getInput("Would you like to change the spreadsheet to work on? (Y = yes, N = no)\n")
                if ChangeActiveWorksheet == "Y":
                        ActiveWorksheet = input("Which sheet would you like to work on?\n").capitalize()
                        changeSheet(ActiveWorksheet)

                ActiveWorksheet = File.active #Needs to be here so that the variable stays as a worksheet object and not change to a str when we change worksheets.
                emptyCells(ActiveWorksheet)
                fillEmptyCell(ActiveWorksheet)
            
            case "C":
                ActiveWorksheet = File.active
                ChangeActiveWorksheet = getInput(f"You are on {ActiveWorksheet}.\nWould you like to change your worksheet? (Y = yes, N = no)\n")
                if ChangeActiveWorksheet == "Y":
                        ActiveWorksheet = input("What sheet do you want to change to?\n").capitalize()
                        changeSheet(ActiveWorksheet)
                else:
                    print(f"Sheet has not been changed. Currently still on {ActiveWorksheet}.")

                CellOption = getInput("Do you want to delete a cell? (Y = yes, N = no)\n")
                match CellOption:
                    case "Y":
                        CellValue = getInput("Enter the cell you want to delete the value of: ")
                        Confirmation = getInput(f"Are you sure you want to delete {ActiveWorksheet[CellValue].value}? (Y = yes, N = no)\n")
                        if (Confirmation == "Y"):
                            ActiveWorksheet[CellValue].value = None
                            File.save(Path)
                    case "N":
                        CellValue = getInput("Enter the cell you want to change the value of: ")
                        print(f"Value at this cell is: {ActiveWorksheet[CellValue].value}")
                        ChangeCellValue = getInput("Enter what you want to put in this cell: ")
                        ActiveWorksheet[CellValue].value = ChangeCellValue
                        print(f"{ActiveWorksheet[CellValue]} has been changed to {ChangeCellValue}")
                        File.save(Path)

def changeSheet(ActiveWorksheet):
    File.active = File[ActiveWorksheet]
    print(f"Current sheet has been changed to: {ActiveWorksheet}")
    return ActiveWorksheet

def cardFinder():
    invalid_input = True
    while True:
        codeORname = getInput("Press \"C\" if you want to find card by its code or \"N\" if you want to search by its name.\nPress \"E\" to exit.\n\nNone of the inputs are case sensitive.\n")
        
        match codeORname:
            case "C" | "N":
                find_card(codeORname)
                invalid_input = False
                findAnotherCard()
                return invalid_input
            case "E":
                print("Exiting search.\n")
                break
            case _:
                print("The input is invalid, please enter \"C\", \"N\" or \"E\".")

def findAnotherCard():
    choice = True
    while True:
        choice = getInput("\nDo you want to find another card? (Y = yes, N = no)\n")
        match choice:
            case "Y":
                choice = True
                cardFinder()
                return choice
            case "N":
                choice = False
                print("Exiting search.\n")
                return choice

#Entry
print("Welcome to the FFTCG spreadsheet searcher!\n\nNote: This spreadsheet doesn't contain cards from opus 7 or 14 because all of the Commons and Rares were obtained from these sets.")
inputChoice()