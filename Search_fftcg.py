from Find_Card import find_card
from path import File

def inputChoice():
    userContinue = True
    while userContinue == True:
        option = input("\nPress one of the following keys to execute their respective functions.\n\nOptions:\n\"E\" to exit the program.\n\"S\" to search for a card.\n\n").upper()
        if (option == "E"):
            print("The program will now close.\nThank you for using this Python script!")
            break
        if (option == "S"):
            cardFinder()

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
        choice = input("\nDo you want to find another card? (Y = yes, N = no)\n").upper()
        match choice:
            case "Y":
                choice = True
                cardFinder()
                return choice
            case "N":
                choice = False
                print("Exiting search.\n")
                return choice

#Variables / Entry
print("Welcome to the FFTCG spreadsheet searcher!\n\nNote: This spreadsheet doesn't contain cards from opus 7 or 14 because all of the Commons and Rares were obtained from these sets.")
inputChoice()
