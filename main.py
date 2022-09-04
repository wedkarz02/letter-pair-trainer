
import os
import sys
import time
import random

try:
    import openpyxl
except:
    print("\033[91m[ERROR]\033[0m: openpyxl not installed.\n")
    os._exit(1)

try:
    if os.path.exists(sys.argv[1]):
        path = sys.argv[1]
        if not path.endswith(".xlsx"):
            print("\033[91m[ERROR]\033[0m: Invalid file privided. See README.md for more information.\n")
            os._exit(1)
    else:
        print(f"\033[91m[ERROR]\033[0m: File '{sys.argv[1]}' does not exist.\n")
        os._exit(1)
except:
    print(f"\033[91m[ERROR]\033[0m: Invalid arguments provided. See README.md for more information.\n")
    os._exit(1)

workbook = openpyxl.load_workbook(path)

try:
    sheet = workbook["Letter Pairs"]
except:
    print("\033[91m[ERROR]\033[0m: No sheet named 'Letter Pairs' found. See README.md for more information.\n")
    os._exit(1)


def print_display(letter_pair_value, lead_letter, follow_letter, parity=False):
    if parity:
        print(f"Parity {lead_letter}: {letter_pair_value}")
    else:
        print(f"{lead_letter}{follow_letter}: {letter_pair_value}")
