
import os
import sys
import time
import random

try:
    import openpyxl
except:
    print("\033[91m[ERROR]\033[0m: openpyxl not installed.")
    os._exit()

if os.path.exists(sys.argv[1]):
    path = sys.argv[1]
    if not path.endswith(".xlsx"):
        print("\033[91m[ERROR]\033[0m: Invalid file privided. See README.md for more information.")
        os._exit()
else:
    print(f"\033[91m[ERROR]\033[0m: File '{sys.argv[1]}' does not exist.")
    os._exit()

workbook = openpyxl.load_workbook(path)
