from oauth2client.service_account import ServiceAccountCredentials as SAC
import os
import numpy as np
from scrapercode import *
from colorama import init, Fore


init(autoreset=True)
while True:
    try:
        setCoin = np.load("data.npz", allow_pickle=True)['coinName'].tolist()
        print("執行目錄:")
        print("1.執行儲存資料")
        print("2.設定抓取資料")
        print("3.設定google試算表")
        print("4.退出")
        anwser = input("\n請輸入執行動作：").strip()
        if anwser == "1":
            saveData()
        elif anwser == "2":  # 設定菜單
            settingCoinUrl()
        elif anwser == "3":
            settingGoogle()
        elif anwser == "4":
            print("\n---------退出...---------")
            os.system("pause")
            break

        else:
            print(Fore.LIGHTRED_EX+"\n--------請輸入數字來執行動作！--------\n")
    except FileNotFoundError:
        coinName = []
        googleUrl = [["", ""]]
        np.savez("data", coinName=coinName, googleUrl=googleUrl)
        print("創建data.npz檔案")
