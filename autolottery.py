
from selenium import webdriver
import time
from scrapy.selector import Selector
import pandas as pd
from datetime import datetime, timedelta
import os
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import sys


def fetch_data():
    chrome_options = Options()

    download_dir = r'C:\Users\mo_ya\Downloads'
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
    })

    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(chrome_options=chrome_options)

    driver.command_executor._commands["send_command"] = (
        "POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {
        'behavior': 'allow', 'downloadPath': download_dir}}
    command_result = driver.execute("send_command", params)

    #     # 撈取檔案
    driver.get('http://cell2.webgene.com.tw/realleaf-admin/index.php')
    time.sleep(1)
    driver.find_element_by_id("account").click()
    driver.find_element_by_id("account").clear()
    driver.find_element_by_id("account").send_keys("webgene")
    driver.find_element_by_id("password").click()
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("password").send_keys("webgene1234")
    driver.find_element_by_xpath("//button[@type='submit']").click()
    driver.find_element_by_name("table_length").click()
    Select(driver.find_element_by_name("table_length")
           ).select_by_visible_text("All")
    driver.find_element_by_name("table_length").click()

    while driver.find_element_by_id('table_processing').is_displayed():
        time.sleep(5)
    else:
        print("continue")

    driver.find_element_by_xpath(
        "//div[@id='table_wrapper']/div[2]/button/span").click()

    filePath = r"C:\Users\mo_ya\Downloads\總覽.xlsx"
    while os.path.exists(filePath) == False:
        print("wait")
        time.sleep(5)

    print("go")


fetch_data()

filePath = r"C:\Users\mo_ya\Downloads\總覽.xlsx"
leads = pd.read_excel(filePath, skiprows=1)

if leads["序號"].count() < 15:
    print("檔案沒有下載成功啦")
    sys.exit(0)
else:
    print("檔案輸入完成")

print("開始資料整理")
leads["時間"] = pd.to_datetime(leads["時間"])
d = datetime.today().strftime("%Y%m%d")
yd = datetime.today() - timedelta(days=1)
date_leads = leads[(leads["時間"] < d) & (
    leads["時間"] > '20200408')]

# 過濾得過獎的人
totalwinners = pd.read_excel(
    r'G:\我的雲端硬碟\共用資料\網站視覺\原萃\天天開獎\總得獎名單.xlsx', index_col=0)
duplicated_leads = date_leads.append(totalwinners)
duplicated_leads.drop_duplicates("序號", keep=False, inplace=True)
duplicated_leads.drop_duplicates("名稱", keep=False, inplace=True)
print("序號是否有重複：{}", duplicated_leads.duplicated(["序號"]).any())
print("名稱是否有重複：{}", duplicated_leads.duplicated(["名稱"]).any())

# 抽獎並輸出本日得獎者
lucky_leads = duplicated_leads.sample(11)
lucky_leads["得獎項目"] = ["頭獎", "二獎", "二獎", "二獎",
                       "二獎", "二獎", "二獎", "二獎", "二獎", "二獎", "二獎"]
writer = pd.ExcelWriter('得獎名單{}_{}_{}.xlsx'.format(
    yd.year, yd.month, yd.day), engine='openpyxl')
lucky_leads.to_excel(writer, sheet_name='Sheet1')
writer.save()
print("輸出本日得獎名單")

# 結合總得獎者名單
total_lucky_leads = totalwinners.append(lucky_leads)
writer = pd.ExcelWriter(
    r'G:\我的雲端硬碟\共用資料\網站視覺\原萃\天天開獎\總得獎名單.xlsx', engine='openpyxl')
total_lucky_leads.to_excel(writer, sheet_name='Sheet1')
writer.save()
print("輸出總得獎名單完成")

os.rename(filePath, r"C:\Users\mo_ya\Downloads\總覽_" +
          str(d) + ".xlsx")
print("總覽檔案改名，今日抽獎完成")
print(lucky_leads)

inputtext = input("按任意鍵結束...")
print(inputtext)
