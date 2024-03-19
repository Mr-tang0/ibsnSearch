import random
import re
import time

import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
import pandas as pd
from openpyxl import Workbook
import os

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


def jdFindBook(book, i=None):
    print(book)
    foundBook = {
        'name': 'null',
        'publish': 'null',
        'time': 'null'
    }
    url = f"https://search.jd.com/Search?keyword={book}"

    url = 'https://item.jd.com/10069415089009.html'
    # import json
    # with open('cookies.json', 'r') as f:
    #     cookies = json.load(f)
    #     driver.add_cookie(cookies)
    #
    # driver.get(url)
    #
    # # cookies = driver.get_cookies()
    # # print(cookies)
    # #
    # # with open('cookies.json', 'w') as f:
    # #     json.dump(cookies, f)

    time.sleep(3)
    a = driver.current_url
    rendered_page = driver.page_source

    soup = BeautifulSoup(rendered_page, "html.parser")
    itemList = soup.find_all(class_="gl-warp clearfix")

    itemList = itemList[0].contents
    for it in itemList:
        sku = re.findall('p-promo-flag">{.*?}<', it)
        if not sku == "广告":
            sku = re.findall('data-sku="{.*?}"', it)
            if not sku == "":
                url = f"https://item.jd.com/{sku}.html"
                break
    time.sleep(1)
    driver.get(url)
    rendered_page = driver.page_source

    isbn = re.findall('<ISBN：', rendered_page)

    soup = BeautifulSoup(rendered_page, "html.parser")

    # sku = re.findall('data-sku="{.*}"', t)
    sku = soup.contents
    print(sku)
    return foundBook


if __name__ == "__main__":
    p = 'test.xlsx'
    booksList = getBookName(p)
    i = 0
    for item in booksList:
        found = jdFindBook(item.get('name'), i)
        i = 1
