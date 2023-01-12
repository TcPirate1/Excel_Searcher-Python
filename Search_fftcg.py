from Find_Card import find_card
from CellFormulas import emptyCells, fillEmptyCell
from path import File, Path

def inputChoice():
    userContinue = True
    while True:
        option = input("\nPress one of the following keys to execute their respective functions.\n\nOptions:\n\"E\" to exit the program.\n\"S\" to search for a card.\n\"A\" to add a new worksheet to the workbook.\n\"W\" to work on a worksheet.\n\"V\" to view current selected worksheet.\n\n\"W\" and \"V\" will show empty cells in the \"Code\" column.\n\"C\" to change the value of a cell.\n\n").upper()
        if (option == "E"):
            print("The program will now close.\nThank you for using this Python script!")
            break
        if (option == "S"):
            cardFinder()
            print(userContinue)
        if (option == "A"):
            new_worksheet = input("What do you want to name this worksheet?\n\nNote: If you don't enter anything it will default to sheet + number.\n")
            File.create_sheet(f"{new_worksheet}")
            File.save(Path)
            print("Worksheet created.")
        if (option == "V"):
            currentWorksheet = File.active
            NoCellValue = input(f"Current sheet is: {currentWorksheet}\n\nWould you like to see the empty cells in {currentWorksheet}?\n").upper()
            if (NoCellValue == "Y"):
                emptyCells(currentWorksheet)
        if (option == "W"):
            print(f"Current sheet is: {File.active}")
            ChangeActiveWorksheet = input("Would you like to change the spreadsheet to work on? (Y = yes, N = no)\n").upper()
            match ChangeActiveWorksheet:
                case "Y":
                    ActiveWorksheet = input("Which sheet would you like to work on?\n").capitalize()
                    changeSheet(ActiveWorksheet)

            ActiveWorksheet = File.active #Needs to be here so that the variable stays as a worksheet object and not change to a str when we change worksheets.
            emptyCells(ActiveWorksheet)
            fillEmptyCell(ActiveWorksheet)
            
        if (option == "C"):
            ActiveWorksheet = File.active
            ChangeActiveWorksheet = input(f"You are on {ActiveWorksheet}.\nWould you like to change your worksheet? ((Y = yes, N = no)\n")
            match ChangeActiveWorksheet:
                case "Y":
                    ActiveWorksheet = input("What sheet do you want to change to?\n").capitalize()
                    changeSheet(ActiveWorksheet)

            CellOption = input("Do you want to delete a cell? ((Y = yes, N = no)\n")
            match CellOption:
                case "Y":
                    CellValue = input("Enter the cell you want to change the value of: ")
                    Confirmation = input("Are you sure you want to delete this cell? ((Y = yes, N = no)\n")
                    if (Confirmation == "Y"):
                        ActiveWorksheet[CellValue].value = None
                        # File.save(Path)
                case "N":
                    CellValue = input("Enter the cell you want to change the value of: ")
                    print(f"Value at this cell is: {ActiveWorksheet[CellValue].value}")
                    ChangeCellValue = input("Enter what you want to put in this cell: ")
                    ActiveWorksheet[CellValue].value = ChangeCellValue
                    print(f"{ActiveWorksheet[CellValue]} has been changed to {ChangeCellValue}")
                    # File.save(Path)

def changeSheet(ActiveWorksheet):
    File.active = File[ActiveWorksheet]
    print(f"Current sheet has been changed to: {ActiveWorksheet}")
    return ActiveWorksheet

def cardFinder():
    invalid_input = True
    while True:
        codeORname = input("Press \"C\" if you want to find card by its code or \"N\" if you want to search by its name.\nPress \"E\" to exit.\n\nNone of the inputs are case sensitive.\n").upper()
        
        match codeORname:
            case "C" | "N":
                find_card(codeORname)
                invalid_input == False
                findAnotherCard()
                break
            case "E":
                print("Exiting search.\n")
                break
            case _:
                print("The input is invalid, please enter \"C\", \"N\" or \"E\".")
                break

def findAnotherCard():
    choice = True
    while True:
        choice = input("Do you want to find another card? (Y = yes, N = no)\n").upper()
        match choice:
            case "Y":
                choice == True
                cardFinder()
                break
            case "N":
                choice == False
                print("Exiting search.\n")
                break

#Variables / Entry
print("Welcome to the FFTCG spreadsheet searcher!\n\nNote: This spreadsheet doesn't contain cards from opus 7 or 14 because all of the Commons and Rares were obtained from these sets.")
inputChoice()
