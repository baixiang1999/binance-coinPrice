from bs4 import BeautifulSoup
from abc import ABC, abstractmethod
from gspread.exceptions import WorksheetNotFound
from colorama import init, Fore
import requests
import json
import gspread
import numpy as np
from oauth2client.service_account import ServiceAccountCredentials as SAC
import time


def saveData():
    try:
        coinName = np.load("data.npz", allow_pickle=True)['coinName'].tolist()
        googleUrlList = np.load("data.npz", allow_pickle=True)[
            'googleUrl'].tolist()
        Json = googleUrlList[0][0]  # Json  表單金鑰
        Url = ["https://spreadsheets.google.com/feeds"]
        Connect = SAC.from_json_keyfile_name(Json, Url)
        GoogleSheets = gspread.authorize(Connect)

        try:
            Sheet = GoogleSheets.open_by_key(
                googleUrlList[0][1])  # 試算表代號
            if coinName == []:
                print(Fore.LIGHTRED_EX+"\n未抓到資料!請再確認是否已設定要抓取的虛擬幣!\n")
            else:
                scraper(Sheet)
        except:
            print(Fore.LIGHTRED_EX + "\n試算表網址尚未輸入或輸入錯誤!\n")
    except:
        print(Fore.LIGHTRED_EX + "\n未找到對應金鑰:"+googleUrlList[0][0]+"\n")


def scraper(Sheet):
    coinName = np.load("data.npz", allow_pickle=True)['coinName'].tolist()
    try:  # 檢查是否有虛擬幣價格表單
        Sheets = Sheet.worksheet("虛擬幣價格")
    except WorksheetNotFound:  # 如果回報找不到表單的錯誤 執行以下
        print("未檢測到工作表，新增工作表中...")
        Sheet.add_worksheet(title="虛擬幣價格", rows=100, cols=100)  # 新增表單
        Sheets = Sheet.worksheet("虛擬幣價格",)
        print(Fore.LIGHTRED_EX+"成功新增工作表")
        Sheets.append_row(["-"], table_range=f"R1C1")

    timeLen = len(Sheets.col_values(1))
    localTime = time.localtime()  # 取得系統時間
    TimeResult = time.strftime("%Y-%m-%d %H:%M", localTime)  # 時間轉換為想要的格式
    timeDatas = [TimeResult]
    Sheets.append_row(
        timeDatas, table_range=f"R{timeLen+1}C1")

    for x in range(len(coinName)):
        try:
            print("抓取" + coinName[x][0] + "價格中...")
            Url = "https://www.binance.com/bapi/asset/v2/public/asset-service/product/get-product-by-symbol?symbol=" + \
                coinName[x][0] + "BUSD"
            priceReqs = requests.get(Url)
            pricereqsjson = json.loads(priceReqs.text)
            priceData = pricereqsjson['data']['c']

            try:  # 檢查第一行是否有對應幣值名稱
                cell = Sheets.findall(coinName[x][0], in_row=1)[0]
                Sheets.update(f"R{timeLen+1}C"+"%s" % (cell.col), priceData)
                print(Fore.LIGHTRED_EX+"---------成功新增" +
                      coinName[x][0]+"幣資料---------\n")
            except IndexError:  # 如果找不到在第一行空白處新增虛擬幣名稱
                nameDatas = coinName[x][0]
                try:    # 找是否有空白處
                    cell = Sheets.find("", in_row=1)
                    Location = "R%sC%s" % (cell.row, cell.col)
                    Sheets.update(Location, nameDatas)
                    Sheets.update("R%sC%s" % (timeLen+1, cell.col), priceData)
                    print(Fore.LIGHTRED_EX+"---------成功新增" +
                          coinName[x][0]+"幣資料---------\n")
                except:  # 沒有空白的話就
                    Sheets.row_values(1)
                    Sheets.update(
                        f'R1C{len(Sheets.row_values(1))+1}', nameDatas)
                    Sheets.update(
                        f'R{timeLen+1}C{len(Sheets.row_values(1))}', priceData)
                    print(Fore.LIGHTRED_EX+"---------成功新增" +
                          coinName[x][0]+"幣資料---------\n")

        except:
            print(Fore.LIGHTRED_EX+"未抓到 " + coinName[x][0] + " 幣的價格...")


def settingCoinUrl():  # 設定菜單
    while True:
        coinName = np.load("data.npz", allow_pickle=True)['coinName'].tolist()
        googleUrl = np.load("data.npz", allow_pickle=True)[
            'googleUrl'].tolist()
        print("\n新增/刪除目前設定")
        print("1.新增")
        print("2.刪除")
        print("3.顯示目前資料")
        print("4.回目錄\n")
        addDele = input("請輸入執行動作：").strip()

        if addDele == "1":
            print("\n新增")
            print("請輸入虛擬幣英文簡寫 ex:BTC")
            addName = str.upper(input("請輸入："))
            coinName.append([addName])
            np.savez("data", coinName=coinName, googleUrl=googleUrl)
            print(Fore.LIGHTRED_EX+"新增　" + addName + " 成功！")

        elif addDele == "2":
            try:  # 嘗試刪除
                print("\n刪除")
                delNum = int(input("請輸入要刪除的項目編號："))
                delName = str(coinName[delNum-1][0])
                del coinName[delNum-1]  # 刪除資料
                print(Fore.LIGHTRED_EX+"成功刪除：" + delName)
                np.savez("data", coinName=coinName, googleUrl=googleUrl)
            except:  # 發生錯誤
                print(Fore.LIGHTRED_EX+"\n編號輸入錯誤!")

        elif addDele == "3":
            print("\n目前已設定抓取資料為以下:")
            print("---------------------------")
            if coinName == [] or coinName == [None]:
                print(Fore.LIGHTRED_EX+"目前無資料")
            else:
                for i in range(len(coinName)):
                    print(str(i+1)+". "+coinName[i][0])
            print("---------------------------")

        elif addDele == "4":
            break
        else:
            print(Fore.LIGHTRED_EX+"\n--------請輸入數字來執行動作！--------\n")


def settingGoogle():  # google設定菜單
    while True:
        googleUrlList = np.load("data.npz", allow_pickle=True)[
            'googleUrl'].tolist()
        coinName = np.load("data.npz", allow_pickle=True)['coinName'].tolist()
        print("\n設定google試算表資料")
        print("1.設定google金鑰")
        print("2.設定google試算表網址")
        print("3.顯示目前設定")
        print("4.回目錄\n")
        googleSettingMenu = input("請輸入執行動作：").strip()
        if googleSettingMenu == "1":
            print("\n設定google金鑰")
            googleGoldKey = input("請輸入'完整'金鑰檔案名稱(需包含.json) : ")
            googleUrlList[0][0] = googleGoldKey
            np.savez("data", coinName=coinName, googleUrl=googleUrlList)
            print(Fore.LIGHTRED_EX+"設定金鑰成功！")
        elif googleSettingMenu == "2":
            print("\n設定google試算表網址")

            googleUrl = input("請輸入試算表'完整'網址：")
            # 整理網址出要用的部分
            googleUrl = googleUrl[googleUrl.rfind(
                "spreadsheets/d/")+15:googleUrl.rfind("edit")-1]

            googleUrlList[0][1] = googleUrl
            np.savez("data", coinName=coinName, googleUrl=googleUrlList)
            print(Fore.LIGHTRED_EX+"設定試算表網址成功！")
        elif googleSettingMenu == "3":
            print("\n目前googole試算表設定如下:")
            print("---------------------------")
            if googleUrlList[0][0] == "":
                print(Fore.LIGHTRED_EX+"尚未設定金鑰")
            else:
                print("金鑰檔名: "+googleUrlList[0][0])

            if googleUrlList[0][1] == "":
                print(Fore.LIGHTRED_EX+"尚未設定試算表網址")
            else:
                print("試算表網址: "+googleUrlList[0][1])
            print("---------------------------")
        elif googleSettingMenu == "4":
            break
        else:
            print(Fore.LIGHTRED_EX+"\n--------請輸入數字來執行動作！--------\n")
