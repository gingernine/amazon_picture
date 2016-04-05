#! usr/env/bin python3
# coding: utf-8

from urllib.request import build_opener
from urllib.parse import urlencode
from PIL import Image
from io import BytesIO
from math import ceil, sqrt
from datetime import datetime
import xlrd
import xlwt
import re
import os


class ExcelIO(object):
    """class handling Excel book I/O"""
    def __init__(self, bookname, sheetname):
        self.bookname=bookname
        self.sheetname=sheetname
        self.create_newbook()

    def create_newbook(self):
        """if file not exists, ceate new book"""
        if not os.path.exists(self.bookname):
            book=xlwt.Workbook()
            sheet=book.add_sheet(self.sheetname)
            book.save(self.bookname)

    def read_book(self):
        readbook=xlrd.open_workbook(self.bookname)
        return readbook

    def write_book(self, data={}):
        readbook=self.read_book()
        work_book=xlwt.Workbook()
        work_sheet=work_book.add_sheet(self.sheetname)
        read_sheet=readbook.sheet_by_name(self.sheetname)
        nrows, ncols = read_sheet.nrows, read_sheet.ncols
        for row in range(nrows):#シートが既にできていれば上書きします
            for col in range(ncols):
                work_sheet.write(row, col, read_sheet.cell_value(row, col))
        colnames={} #列名と列番号を取得します
        if nrows: #すでにbookにデータが登記されている場合
            i=1
            while 1:
                try:
                    colnames[read_sheet.cell_value(0, i)]=i
                    i+=1
                except: break
        else: #bookにデータが未入力の場合
            work_sheet.write(0, 0, '更新日付')
            for i, colname in enumerate(data):
                work_sheet.write(0, i+1, colname)
                colnames[colname]=i+1
            nrows=1
        work_sheet.write(nrows, 0, str(datetime.today()))
        for col in colnames:
            work_sheet.write(nrows, colnames[col], data[col])
        work_book.save(self.bookname)

opener=build_opener()

def get_contents(url):
    """get HTML scripts"""
    try:
        with opener.open(url) as con:
            contents=con.read()
        return contents
    except IOError:
        print('このurlを開くことが出来ません: %s'%url)


def download_img(asin, dirpath):
    """download images from amazon site"""
    url='http://www.amazon.co.jp/s?__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&url=search-alias%3Daps&field-keywords='+asin
    contents=get_contents(url).decode('utf-8') #ASINでの検索結果のページのhtmlスクリプトを取得します
    if '検索に一致する商品はありませんでした。' in contents:
        print('検索に一致する商品はありませんでした。asinが正しいかを確認してください')
        return
    targetlink_pat=re.compile('class="a-link-normal a-text-normal" target="_blank" href="(.*?)"')
    targetlink_url=targetlink_pat.search(contents).group(1) #上で取得したスクリプトから対称のページのURLを取得します
    contents=get_contents(targetlink_url).decode('utf-8')
    img_url_list=re.findall('{"hiRes":.*?"large":"(.*?)".*?}',contents) #画像のURLを取得します
    if not img_url_list:
        print('対象商品の画像がありません')
        return
    img_num=len(img_url_list) #画像枚数を取得します
    register=ExcelIO('C:\\Downloads\\register.xls', 'new')
    register.write_book(data={'ASIN': asin, '画像の枚数': img_num, '保存先': dirpath})
    for i,url in enumerate(img_url_list):
        resize_img(i, url, asin, dirpath)


def resize_img(i, url, asin, dirpath):
    """if img_size<500, resize img"""
    imgdata=get_contents(url)
    img=Image.open(BytesIO(imgdata))
    img_size=img.size
    if img_size[0]*img_size[1] >= 500:
        with open(dirpath+'\\'+asin+'-%d'%(i+1)+'.jpg', 'wb') as wimg:
            wimg.write(imgdata)
        return
    aspect_ratio=img_size[1]/img_size[0]
    width=ceil(sqrt(500/aspect_ratio))
    height=ceil(width*aspect_ratio)
    resized_img=img.resize((width, height))
    resized_img.save(dirpath+'\\'+asin+'-%d'%(i+1)+'.jpg','JPEG',quolity=100, optumize=True)


if __name__=='__main__':
    asin=input('ASINを入力してください ---> ')
    folder=input('保存先のフォルダ名を入力してください ---> ')
    dirpath='C:\\Downloads\\'+folder
    if not os.path.exists(dirpath):
        os.makedirs(dirpath)
    download_img(asin, dirpath)
