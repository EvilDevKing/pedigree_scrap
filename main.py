from colorama import Fore, init
import subprocess as sp

import pdfscript as pdf
import gethorseage as age

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'Brittany Holy Script'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

if __name__ == '__main__':
    init() # initialize a colorama usage
    
    print(f"{Fore.GREEN}[1] Start buckle pdf script")
    print(f"[2] Start horse age script{Fore.RESET}")
    print("\n")
    num = input(f"{Fore.CYAN}[x] Select a number to execute script: {Fore.RESET}")
    sheetId = input(f"{Fore.CYAN}Input a specific google sheet id what you want: {Fore.RESET}")
    if sheetId.strip() == "":
        print("Sorry, you didn't provide your sheet id.")
    else:
        if num == "1":
            pdf.run(sheetId)
        elif num == "2":
            sheet_name = input(f"{Fore.CYAN}Input a specific sheet name for horse age: {Fore.RESET}")
            if sheet_name.strip() == "":
                print("Sorry, you didn't provide a sheet name.")
            else:
                age.run(sheetId, sheet_name)
    input("Press any key to exit...")