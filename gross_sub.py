import os
import loguru
from genericpath import exists
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font
import requests
import time
import datetime
import random
import pandas as pd


def time_title():
    print('\n===================================')
    print(' Time : '+str(datetime.datetime.today()))
    print('===================================')


def rand_on():
    delay_lst = [0.1, 0.2, 0.3, 0.4 ,0.5]
    delay = random.choice(delay_lst)
    time.sleep(delay)


def get_stock_datetime():
    reqs_date = requests.get("https://dj.mybank.com.tw/z/zc/zcl/zcl_2330.djhtm")

    if reqs_date.status_code != 200:
        loguru.logger.error('REQS: status code is not 200')
    loguru.logger.success('REQS: success')

    soup_date = BeautifulSoup(reqs_date.text, 'html.parser')
    date_tmp = (((soup_date.find('table', {'class': 't01'})).find_all('tr')[7]).find_all('td')[0]).text
    return date_tmp


def xls_wb_on(path_xls):
    return openpyxl.load_workbook(path_xls) if os.path.exists(path_xls) else openpyxl.Workbook()


def xls_st_on(obj,flg,st_name,idx):
    for stn in obj.sheetnames: flg+=1 if stn == st_name else +0
    sheet = obj[st_name] if flg == 1 else obj.create_sheet(st_name,idx)
    return sheet


def get_stock_urls(Stock_Num):
    urls = []
    #urls.append(f'http://kgieworld.moneydj.com/ZXQW/zc/zca/zca_{Stock_Num}.djhtm')
    #urls.append(f'https://kgieworld.moneydj.com/z/zc/zca/zca_{Stock_Num}.djhtm')
    #urls.append(f'http://jsjustweb.jihsun.com.tw/z/zc/zca/zca_{Stock_Num}.djhtm')
    urls.append(f'https://dj.mybank.com.tw/z/zc/zca/zca_{Stock_Num}.djhtm')
    return urls


@loguru.logger.catch
def get_reqs_data(urls):
    return [ requests.get(link) for link in urls ]
    '''
    for link in urls:
        reqs = requests.get(link)
        if reqs.status_code != 200:
            loguru.logger.error('REQS: status code is not 200')
        loguru.logger.success('REQS: success')
    print(type(reqs))
    return reqs
    '''


@loguru.logger.catch
def parse_stock_data(reqs):
    for r in reqs:
        soup = BeautifulSoup(r.text,'html.parser')
        blocks = soup.find_all('table', {'class': 't01'})
        for blk in blocks:
            dat_p = float((((blk.find_all('tr')[1]).find_all('td')[7].text)).replace(',',''))
            dat_v = float((((blk.find_all('tr')[3]).find_all('td')[7].text)).replace(',',''))
            dat_c = float((((blk.find_all('tr')[13]).find_all('td')[1].text)).replace(',',''))
            dat_t = round((dat_v/dat_c/100),2)
    return dat_p,dat_v,dat_t


@loguru.logger.catch
def parse_stock_data_yahoo(reqs): # for checking the price
    for r in reqs:
        soup = BeautifulSoup(r.text,'html.parser')
        if   soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c) C($c-trend-up)'}) is not None :
            Yahoo_Price = (soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c) C($c-trend-up)'})).text
        elif soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c) C($c-trend-down)'}) is not None:
            Yahoo_Price = (soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c) C($c-trend-down)'})).text
        elif soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c)'}) is not None:
            Yahoo_Price = (soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c)'})).text
        elif soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px)'}) is not None:
            Yahoo_Price = (soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px)'})).text
        elif soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) C(#fff) Px(6px) Py(2px) Bdrs(4px) Bgc($c-trend-up)'}) is not None:
            Yahoo_Price = (soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) C(#fff) Px(6px) Py(2px) Bdrs(4px) Bgc($c-trend-up)'})).text
        elif soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) C(#fff) Px(6px) Py(2px) Bdrs(4px) Bgc($c-trend-down)'}) is not None:
            Yahoo_Price = (soup.find('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) C(#fff) Px(6px) Py(2px) Bdrs(4px) Bgc($c-trend-down)'})).text
        #print(str(Yahoo_Price))
    return Yahoo_Price


@loguru.logger.catch
def cal_avg_price(obj,lst,row):
    sum_tmp=[]
    avg_tmp=[]
    for i in range(len(lst)): sum_tmp.append(0)
    for i in range(len(lst)): avg_tmp.append(0)
    for cnt in range(len(lst)):
        for clm in range(obj.max_column,obj.max_column-lst[cnt],-1): sum_tmp[cnt]+=(obj.cell(row=row, column=clm)).value
        avg_tmp[cnt]=round(sum_tmp[cnt]/lst[cnt],2)
    return avg_tmp


@loguru.logger.catch
def cal_increase_rate(obj,lst,row):
    pri_tmp=[]
    rat_tmp=[]
    for i in range(len(lst)): pri_tmp.append(0)
    for i in range(len(lst)): rat_tmp.append(0)
    today_price = (obj.cell(row=row, column=obj.max_column)).value
    for cnt in range(len(lst)):
        pri_tmp[cnt]=(obj.cell(row=row, column=obj.max_column-lst[cnt])).value
        rat_tmp[cnt]=round(float((today_price-pri_tmp[cnt])/pri_tmp[cnt]*100),2) if pri_tmp[cnt] != 0 else 0
    return rat_tmp


@loguru.logger.catch
def cal_slope_rate(obj,r):
    d1 = float((obj.cell(row=r, column=obj.max_column-0)).value)
    d2 = float((obj.cell(row=r, column=obj.max_column-5)).value)
    val = round(((d1-d2)/d1)*100/5,4)
    return val


@loguru.logger.catch
def cal_price_position(obj1,obj2,r,nam):
    if nam == '20ma' : nam = '月線_'
    if nam == '60ma' : nam = '季線_'
    tday_price = float((obj1.cell(row=r, column=obj1.max_column)).value)
    line_price = float((obj2.cell(row=r, column=obj2.max_column)).value)
    if   tday_price  > line_price : pos_cmt = nam+'上'
    elif tday_price == line_price : pos_cmt = nam
    elif tday_price  < line_price : pos_cmt = nam+'下'
    return pos_cmt
