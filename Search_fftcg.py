from Find_Card import find_card
from ColumnConverter import num_hash
from EmptyCells import emptyCells
from path import File
import os

def inputChoice():
    userContinue = True
    while True:
        option = input("\nPress one of the following keys to execute their respective functions.\nOptions:\n\"E\" to exit the program.\n\"S\" to search for a card.\n\"A\" to add a new worksheet to the workbook.\n\"W\" to work on a worksheet.\n\"V\" to view current selected worksheet.\n\n").upper()
        if (option == "E"):
            print("The program will now close.\nThank you for using this Python script!")
            break
        if (option == "S"):
            cardFinder()
            doSomethingElse(userContinue)
            break
        if (option == "A"):
            new_worksheet = input("What do you want to name this worksheet?\n\nNote: If you don't enter anything it will default to sheet + number.\n")
            File.create_sheet(f"{new_worksheet}")
            File.save(os.path.join(os.path.expanduser('~'), 'Documents', 'Spreadsheets' , 'FFTCG_automated.xlsx'))
            print("Worksheet created.")
            doSomethingElse(userContinue)
            break
        if (option == "V"):
            currentWorksheet = File.active
            print(f"Current sheet is: {currentWorksheet}")
        if (option == "W"):
            print(f"Current sheet is: {File.active}")
            ChangeActiveWorksheet = input("Would you like to change the spreadsheet to work on? (Y = yes, N = no)\n").upper()
            match ChangeActiveWorksheet:
                case "Y" | "YES":
                    ActiveWorksheet = input("Which sheet would you like to work on?\n").capitalize()
                    File.active = File[ActiveWorksheet]
                    print(f"Current sheet has been changed to: {ActiveWorksheet}")

            ActiveWorksheet = File.active
            emptyCells(ActiveWorksheet)
            # selected_cell = input("Choose the cell you want to write to:\n").upper()
            # change_value = input("Enter what you want to put in this cell:\n").upper()
            # ActiveWorksheet[selected_cell] = change_value
            # print(ActiveWorksheet[selected_cell])


def doSomethingElse(userContinue):
    userinput = input("\nWould you like to do something else? (Y = yes, N = no)\n").upper()
    match userinput:
        case "N" | "NO":
            userContinue == False
        case "Y" | "YES":
            userContinue == True
    return userContinue

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
