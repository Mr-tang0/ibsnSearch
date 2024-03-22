import random
import re
import time
import json
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import os
from datetime import datetime

edge_options = Options()
# edge_options.add_argument('--headless')
# edge_options.add_argument('--disable-gpu')
# edge_options.add_argument('--disable-cookies')
# edge_options.add_argument('--incognito')

ser = Service()
ser.path = 'D:/download/IEDownload/msedgedriver.exe'

driver = webdriver.Edge(service=ser, options=edge_options)


def getBookName(path):
    file = pd.read_excel(path, sheet_name='Sheet1')
    bookList = file.iloc[:, 1]
    publishList = file.iloc[:, 2]
    timeList = file.iloc[:, 3]

    books = []
    for i in range(len(bookList)):
        book = {
            'name': bookList[i],
            'publish': publishList[i],
            'time': timeList[i]
        }
        books.append(book)
    return books


def jdFindBook(book):
    url = f"https://search.jd.com/Search?keyword={book}"
    driver.get(url)
    rendered_page = driver.page_source

    soup = BeautifulSoup(rendered_page, "html.parser")
    itemList = soup.find_all(class_="gl-warp clearfix")
    itemList = itemList[0].contents
    while '\n' in itemList:
        itemList.remove('\n')
    for rangeItem in itemList:
        adFlag = rangeItem.find_all(class_="p-promo-flag")
        if not adFlag:
            sku = rangeItem.attrs.get("data-sku")
            url = f"https://item.jd.com/{sku}.html"
            break
    time.sleep(random.randint(1, 3))
    driver.get(url)
    book_page = driver.page_source

    book_soup = BeautifulSoup(book_page, "html.parser")
    name = book_soup.find_all(class_='sku-name')[0].contents
    if name:
        name = name[0]
        name = name.strip()
    else:
        name = "null"
    details = book_soup.find_all(class_='parameter2 p-parameter-list')

    txt = str(details)
    publish = re.findall('title="(.*?)">出版社', txt)
    if publish:
        publish = publish[0]
        publish = publish.strip()
    else:
        publish = "null"

    ISBN = re.findall('ISBN：(.*?)</li>', txt)
    if ISBN:
        ISBN = ISBN[0]
        ISBN = ISBN.strip()
    else:
        ISBN = 'null'

    TIME = re.findall('出版时间：(.*?)</li>', txt)
    if TIME:
        TIME = TIME[0]
        TIME = TIME.strip()
    else:
        TIME = 'null'

    Book = {
        'name': name,
        'publish': publish,
        'ISBN': ISBN,
        'time': TIME
    }

    return Book


def saveToExcel(old, new, path):
    if not os.path.isfile(path):
        wb = Workbook()
        wb.save(path)

    file = pd.read_excel(path, header=None)

    oldNameLine = []
    oldTimeLine = []
    oldPublishLine = []
    newNameLine = []
    newTimeLine = []
    newPublishLine = []
    newISBNLine = []

    for rand in range(0, len(old), 1):
        oldNameLine.append(old[rand].get('name'))
        oldPublishLine.append(old[rand].get('publish'))
        oldTimeLine.append(old[rand].get('time'))

        newNameLine.append(new[rand].get('name'))
        newTimeLine.append(new[rand].get('time'))
        newPublishLine.append(new[rand].get('publish'))
        newISBNLine.append(new[rand].get('ISBN'))

    column1 = pd.Series(oldNameLine, name='oldNameLine')
    column2 = pd.Series(oldPublishLine, name='oldPublishLine')
    column3 = pd.Series(oldTimeLine, name='oldTimeLine')

    column4 = pd.Series(newNameLine, name='newNameLine')
    column5 = pd.Series(newTimeLine, name='newTimeLine')
    column6 = pd.Series(newPublishLine, name='newPublishLine')
    column7 = pd.Series(newISBNLine, name='newISBNLine')

    if not file.empty:
        file.insert(loc=0, column=column1.name, value=column1)
        file.insert(loc=1, column=column4.name, value=column4)
        file.insert(loc=2, column=column2.name, value=column2)
        file.insert(loc=3, column=column6.name, value=column6)
        file.insert(loc=4, column=column3.name, value=column3)
        file.insert(loc=5, column=column5.name, value=column5)
        file.insert(loc=6, column=column7.name, value=column7)

    else:
        # 初始化一个新的DataFrame，并确保数据按行排列
        file = pd.DataFrame({'oldNameLine': oldNameLine,
                             'newNameLine': newNameLine,

                             'oldPublishLine': oldPublishLine,
                             'newPublishLine': newPublishLine,

                             'oldTimeLine': oldTimeLine,
                             'newTimeLine': newTimeLine,

                             'newISBNLine': newISBNLine,
                             })

    file.to_excel(path, index=False)


if __name__ == "__main__":
    i = 0

    # 导入要查的书
    booksList = getBookName('test.xlsx')

    # 导入cookie
    with open('jdcookies.json', 'r') as f:
        cookies = json.load(f)
    driver.get("https://jd.com")
    time.sleep(random.randint(2, 5))

    for cookie in cookies:
        if 'expiry' in cookie:
            del cookie['expiry']
        driver.add_cookie(cookie)

    # 查找
    foundBooks = []
    oldBooks = []
    for book in booksList:

        time.sleep(random.randint(3, 7))
        foundBook = jdFindBook(book.get('name'))

        foundBooks.append(foundBook)
        print(foundBook)
        oldBooks.append(book)

        i = i + 1
        if i % 10 == 0:
            saveToExcel(old=oldBooks, new=foundBooks, path=str(i) + "_newTest.xlsx")
            print("save:" + str(i) + "newTest.xlsx")
            foundBooks.clear()
            oldBooks.clear()

    driver.quit()
