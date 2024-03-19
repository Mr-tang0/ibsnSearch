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


def saveBookIbsn(path="", isbnL=None, nameL=None):
    # 检查文件是否存在，如果不存在则创建一个新Excel文件（使用openpyxl）
    if not os.path.isfile(path):
        wb = Workbook()
        wb.save(path)

    file = pd.read_excel(path, header=None)

    column1 = pd.Series(isbnL, name='ISBN')
    column2 = pd.Series(nameL, name='Name')

    if not file.empty:
        file.insert(loc=0, column=column1.name, value=column1)
        file.insert(loc=1, column=column2.name, value=column2)
    else:
        # 初始化一个新的DataFrame，并确保数据按行排列
        file = pd.DataFrame({'ISBN': isbnL, 'Name': nameL})

    file.to_excel(path, index=False)
    # with pd.ExcelWriter(path, engine='openpyxl', mode='a') as w:
    #     file.to_excel(w, sheet_name='Sheet1', index=False)
    #     w._save()

    # file.insert(loc=1, column=column.name, value=column)
    # file.to_excel(path, index=False)


class Isbn:
    def __init__(self):
        self.edge_options = Options()
        # self.edge_options.add_argument('--headless')
        # self.edge_options.add_argument('--disable-gpu')
        self.edge_options.add_argument('--disable-cookies')
        self.edge_options.add_argument('--incognito')

        self.ser = Service()
        self.ser.path = 'D:/download/IEDownload/msedgedriver.exe'

        self.driver = webdriver.Edge(service=self.ser, options=self.edge_options)

    def get_book_kongfuzi(self, name=""):
        url = f"https://search.kongfz.com/product_result/?key={name}"

        # url = f"https://search.douban.com/book/subject_search?search_text={name}&cat=1001"
        self.driver.get(url)
        time.sleep(1)
        rendered_page = self.driver.page_source

        soup = BeautifulSoup(rendered_page, "html.parser")
        itemList = soup.find_all(class_="item clearfix")

        books = []
        for item in itemList:
            author = re.findall(r' 作者:  (.*?) {3}出版社', item.text)
            if len(author) > 0:
                author = author[0]
            else:
                author = ""

            publishYear = re.findall(r'出版时间:  (.*?) {3}装帧', item.text)
            if len(publishYear) > 0:
                publishYear = publishYear[0]
            else:
                publishYear = ""

            publish = re.findall(r'出版社 (.*?) ', item.text)

            book = {
                'name': item.attrs.get('itemname'),
                'isbn': item.attrs.get('isbn'),
                'author': author,
                'publish_year': publishYear,
                'publish': publish
            }
            books.append(book)
        # if len(books) != 0:
        #     print("找到书" + name)

        for dictItem in books:
            if dictItem.get('isbn') != "":
                if dictItem.get('name') != "":
                    if dictItem.get('name')[0] == name[0] and dictItem.get('name')[-1] == name[-1]:
                        # if dictItem.get('publish_year')[:3] == year:
                        return dictItem.get('isbn'), dictItem.get('name')
        return 'null', 'null'

    def get_book_dangdang(self, na="", n=None):
        url = f"https://search.dangdang.com/?key={na}&act=input"
        self.driver.get(url)
        time.sleep(3)
        rendered_page = self.driver.page_source

        soup = BeautifulSoup(rendered_page, "html.parser")
        itemList = soup.find_all(class_="line1")
        if len(itemList) == 0:
            return na, "null"

        itemList = itemList[0].contents
        itemList = re.findall('href="(.*?)"', str(itemList))
        if not len(itemList) == 0:
            itemList = itemList[0]
            url = f"https:{itemList}"

            if n == 0:
                time.sleep(15)
            else:
                time.sleep(random.randint(1, 3))

            self.driver.get(url)
            rendered_page = self.driver.page_source
            soup = BeautifulSoup(rendered_page, "html.parser")

            ISBN = soup.find_all(class_="key clearfix")

            if len(ISBN) == 0:
                return na, "null"
            text = ISBN[0].text
            ISBN = re.findall('ISBN：(.*)', text)

            Name = soup.find_all(class_="name_info")
            text2 = str(Name)
            Name = re.findall('title="(.*?)"', text2)
            Name = Name[0]
            print()

            if not len(ISBN) == 0:
                print()
                return Name, ISBN[0][0:13]
            else:
                return na, "null"
        else:
            return na, "null"

    def getBookName(self, path=""):
        file = pd.read_excel(path, sheet_name='Sheet1')
        bookList = file.iloc[:, 1]
        authorList = file.iloc[:, 2]

        books = []
        for i in range(len(bookList)):
            book = {
                'name': bookList[i],
                'author': authorList[i],
            }
            books.append(book)
        return books


if __name__ == "__main__":
    base = Isbn()
    gotBooks = base.getBookName("test.xlsx")

    isbnList = []
    nameList = []

    num = 0
    i = 0

    for book in gotBooks:
        time.sleep(random.randint(2, 5))
        name, isbn = base.get_book_dangdang(na=str(book.get('name'))[:13], n=num)
        print(name, isbn)
        isbnList.append(isbn)
        nameList.append(name)
        num = num + 1
        if num % 25 == 0:
            saveBookIbsn(str(num) + "newTest.xlsx", isbnList, nameList)
            print("save:" + str(num) + "newTest.xlsx")
            isbnList.clear()
            nameList.clear()
    saveBookIbsn("end" + "newTest.xlsx", isbnList)
    print("Ok")
    base.driver.quit()
