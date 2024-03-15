import random
import re
import time

import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
import pandas as pd
import openpyxl


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

            book = {
                'name': item.attrs.get('itemname'),
                'isbn': item.attrs.get('isbn'),
                'author': author,
                'publish_year': publishYear
            }
            books.append(book)
        # if len(books) != 0:
        #     print("找到书" + name)

        for dictItem in books:
            if dictItem.get('isbn') != "":
                if dictItem.get('name') != "":
                    if dictItem.get('name')[0] == name[0] and dictItem.get('name')[-1] == name[-1]:
                        # if dictItem.get('publish_year')[:3] == year:
                        return dictItem.get('isbn')


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

    def saveBookIbsn(self, path="", List=None):
        file = pd.read_excel(path, sheet_name='Sheet1', header=None)
        column = pd.Series(List)
        file.insert(loc=1, column=column.name, value=column)
        file.to_excel(path, index=False)


if __name__ == "__main__":
    base = Isbn()
    gotBooks = base.getBookName("test.xlsx")

    isbnList = []
    for book in gotBooks:
        time.sleep(random.randint(3, 10))
        print(book.get('name'))
        isbn = base.get_book_kongfuzi(name=book.get('name'))
        print(isbn)
        isbnList.append(isbn)
    base.saveBookIbsn("newTest.xlsx", isbnList)
    print("Ok")
    base.driver.quit()
