from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import datetime
from selenium import webdriver
from openpyxl import Workbook,load_workbook
import time


date = datetime.date.today().strftime("%d-%m-%Y")

s = Service(r"D:\Durai\Driver\chromedriver.exe")
driver = webdriver.Chrome(service=s)
#
# l_wb =load_workbook(r"D:\Durai\Scraping\Mobile\Save Data\Poorvilka Files\Mobile item code 22-11-2022.xlsx")
# l_ws = l_wb.active

wb = Workbook()
ws = wb.active
ws.title = "Mobile"

l = 1
# for r in range(1, 70):
for r in range(1, 3):
    print("                  ")
    print(r)

    driver.get(url = "https://www.poorvika.com/s?categories=categories.lvl1%3A%3D%5B%60Computers+%26+Tablets+%3E+Tabs+%26+IPad%60%5D&stock_status=stock_status%3A%3D%5B%60In+Stock%60%5D&page=" + str(r))
    time.sleep(2)

    for box in driver.find_elements(By.CLASS_NAME, "product-card_card__description__2LoI_"):
        name = box.find_element(By.TAG_NAME, "b").text
        print("Name = ",name)
        ws.cell(row=l, column=1).value = name

        # for link in box.find_elements(By.XPATH,'//*[@id="__next"]/div/div[4]/div/div[1]/div[1]/div/div[2]/div[1]/div/div[1]/div[2]/div/span'):
        #     print(link.text)
        #     ws.cell(row=l, column=2).value = link.text

        wb.save("D:\Durai\Scraping\Mobile\Save Data\Poorvilka Files\Mobile 12 " + date + ".xlsx")
        l = l + 1
# text-sm center-content_item_code__3s1nf