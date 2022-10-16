from Find_Card import find_card

def inputChoice():
    codeORname = input("Press \"C\" if you want to find card by its code or \"N\" if you want to search by its name: ").upper()
    if (codeORname == "C"):
        find_card(codeORname)
    if (codeORname == "N"):
        find_card(codeORname)

#Variables / Entry
print("Welcome to the FFTCG spreadsheet searcher!\n")
inputChoice()

choice = True
while True:
    choice = input("Do you want to find another card? (Y = yes, N = no)\n").upper()
    if (choice == "Y"):
        choice == True
        inputChoice()

    if (choice == "N"):
        choice == False
        print("Thank you for using this Python script!")
        break