from Find_Card import find_card

def inputChoice():
    invalid_input = True
    while True:
        codeORname = input("Press \"C\" if you want to find card by its code or \"N\" if you want to search by its name.\nPress \"E\" to exit.\n").upper()
        if (codeORname == "C" or codeORname == "N"):
            find_card(codeORname)
            invalid_input == False
            doAgain()
            break
        if (codeORname == "C" or codeORname == "N" or codeORname == "E"):
            print("The input is invalid, please enter \"C\", \"N\" or \"E\".")
        if (codeORname == "E"):
            break
def doAgain():
    choice = True
    while True:
        choice = input("Do you want to find another card? (Y = yes, N = no)\n").upper()
        if (choice == "Y"):
            choice == True
            inputChoice()
            break
        if (choice == "N"):
            choice == False
            print("Thank you for using this Python script!")
            break

#Variables / Entry
print("Welcome to the FFTCG spreadsheet searcher!\n")
inputChoice()