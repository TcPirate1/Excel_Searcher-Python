from Find_Card import find_card
from EmptyCells import emptyCells
from path import File, Path
import os

def inputChoice():
    userContinue = True
    while True:
        option = input("\nPress one of the following keys to execute their respective functions.\nOptions:\n\"E\" to exit the program.\n\"S\" to search for a card.\n\"A\" to add a new worksheet to the workbook.\n\"W\" to work on a worksheet.\n\"V\" to view current selected worksheet.\n\n\"W\" and \"V\" will show empty cells in the \"Code\" column.\n\"C\" to change the value of a cell.\n\n").upper()
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
            if (NoCellValue == "Y" or "YES"):
                emptyCells(currentWorksheet)
        if (option == "W"):
            print(f"Current sheet is: {File.active}")
            ChangeActiveWorksheet = input("Would you like to change the spreadsheet to work on? (Y = yes, N = no)\n").upper()
            match ChangeActiveWorksheet:
                case "Y" | "YES":
                    ActiveWorksheet = input("Which sheet would you like to work on?\n").capitalize()
                    File.active = File[ActiveWorksheet]
                    print(f"Current sheet has been changed to: {ActiveWorksheet}")

            ActiveWorksheet = File.active #Needs to be here so that the variable stays as a worksheet object and not change to a str when we change worksheets.
            emptyCells(ActiveWorksheet)
            print("Reminder, these cells are under the \"Code\" column!")
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
            # File.save(Path)
        if (option == "C"):
            ActiveWorksheet = File.active
            print("What sheet do you want to change?")
            print("What cell do you want to change?")
            print(f"You are on {ActiveWorksheet}.\nWould you like to change your worksheet? ((Y = yes, N = no)\n")


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
            case "Y" | "YES":
                choice == True
                cardFinder()
                break
            case "N" | "NO":
                choice == False
                print("Exiting search.\n")
                break

#Variables / Entry
print("Welcome to the FFTCG spreadsheet searcher!\n\nNote: This spreadsheet doesn't contain cards from opus 7 or 14 because all of the Commons and Rares were obtained from these sets.")
inputChoice()
