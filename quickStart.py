from oauth2client.service_account import ServiceAccountCredentials as SAC
import os
import numpy as np
from scrapercode import *
from colorama import init, Fore

init(autoreset=True)
try:
    saveData()
except FileNotFoundError:
    print("尚未進行任何設定!請先執行menu.exe進行設定")
print("10秒後自動關閉...")
time.sleep(10)
input()
