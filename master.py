from colorama import init, Fore
from constants import getGoogleService

import master_1d as m1
import master_overall as m2

def checkIsMasterSheetId(sheetId):
    service = getGoogleService("sheets", "v4")
    flag = False
    try:
        sheet_metadata = service.spreadsheets().get(spreadsheetId=sheetId).execute()
        sheets = sheet_metadata.get('sheets', '')
        for sheet in sheets:
            if sheet["properties"]["title"] == "Overall Master crosses" or sheet["properties"]["title"] == "1D crosses":
                flag = True
    except:
        flag = False
    return flag

if __name__ == "__main__":
    init()
    sheetId = input(f"{Fore.CYAN}Enter your master sheet id: {Fore.RESET}")
    if sheetId.strip() != "":
        if checkIsMasterSheetId(sheetId):
            print(f"{Fore.GREEN}1. 1D crosses")
            print(f"2. Overall Master crosses{Fore.RESET}")
            print("")
            opt = input(f"{Fore.MAGENTA}Enter a number: {Fore.RESET}")
            if opt.strip() == "":
                print(f"{Fore.RED}Must enter a number.{Fore.RESET}")
            else:
                try:
                    opt = int(opt.strip())
                    if opt == 1:
                        m1.run(sheetId)
                    else:
                        m2.run(sheetId)
                except:
                    print(f"{Fore.RED}Must enter a digit.{Fore.RESET}")
        else:
            print(f"{Fore.RED}The sheet id is not valid or not for master sheet.{Fore.RESET}")
    else:
        print(f"{Fore.RED}The master sheet id can't be empty.{Fore.RESET}")