from Bank_func import *


folder = "Data"

banks = file_to_banks()
ID = 0
NAME = 1
NUMBER_OF_CARDS = 2

operations = file_to_operations()
BANK = 0
DATE = 1
TIME = 2
CARD = 3
OPERATION = 4
MONEY = 5
BALANCE = 6
CURRENCY = 7

menu = ""
a = True
while menu == "" and a:
    print("\nChoose an option: \n"
          "1 - Show current funds \n"
          "2 - Expenses per month \n"
          "3 - Exit from the program")
    menu = str(input("Your choice: "))
    while menu == "1":
        print("\nYour current funds:")
        summary = 0
        cards = all_cards_in_banks(operations2)
        k = len(cards)
        while k > 0:
            one_card = cards[k-1]
            i = 0
            boolean = True
            while i < len(operations2) and boolean:
                if one_card == operations2[i][CARD]:
                    b = balance(one_card)[0]
                    print(f"{operations2[i][CARD]} ({name_by_number(int(operations2[i][BANK]))}): {b} "
                          f"{operations2[i][CURRENCY]}")
                    boolean = False
                    summary += int(b)
                i += 1
            k -= 1
        print(f"Total: {summary} EUR")
        input("\nPress any key")
        menu = ""
    while menu == "2":
        time = str(input("Enter month and year in the following format MM-YYYY: "))
        index_of = time.find("-")
        if len(time) == len("MM-YYYY") and index_of == 2:
            m, z = balance_in_month(time)
            rsd, length = received_spent_delta(m)
            if length != 0:
                while menu != "":
                    print("\nSelect a credit card: \n")
                    i = 0
                    while i < len(rsd):
                        print(f"{i+1} - {rsd[i][0]} ({name_by_number(number_by_card(rsd[i][0]))})")
                        i += 1
                    print(f"{len(rsd)+1} - Total")
                    print(f"{len(rsd)+2} - Exit to the main menu")
                    selected_card = int(input("\nYour choice: "))
                    cur = operations[1][CURRENCY]
                    if 1 <= selected_card <= len(rsd):
                        s_card = rsd[selected_card-1][0]
                        nbc = number_by_card(s_card)
                        print(f"Report for {date_to_str(time)}, {s_card} ({name_by_number(nbc)})\n"
                              f"Received: {rsd[selected_card-1][1]} {cur}\nSpent: {rsd[selected_card-1][2]} {cur}\n"
                              f"Delta: {rsd[selected_card-1][3]} {cur}")
                        report_choice = str(input("Export a full report to Excel? (y/n): "))
                        if report_choice == 'y':
                            export_to_excel(time, f"{s_card}", rsd)
                            print("Report is created")
                    elif selected_card == len(rsd) + 1:
                        print(f"Report for {date_to_str(time)}, all credit cards\n"
                              f"Received: {rsd_for_all_cards(rsd)[0]} {cur}\nSpent: {rsd_for_all_cards(rsd)[1]} {cur}\n"
                              f"Delta: {rsd_for_all_cards(rsd)[2]} {cur}")
                        report_choice = str(input("Export a full report to Excel? (y/n)"))
                        if report_choice == 'y':
                            export_to_excel(time, "all cards", rsd)
                            print("Report is created")
                    elif selected_card == len(rsd) + 2:
                        menu = ""
            elif length == 0:
                print("\nNo operations in selected month.\nPlease, try again!\n")
        else:
            print("\nIncorrect input. Please, try again\n")
    if menu == "3":
        print("Program shutdown...")
        a = False
    elif menu != "1" and menu != "2" and menu != "3" and menu != "":
        print("\nError! Please, try again\n")
        menu = ''
