import os
from genericpath import exists
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font
import requests
import time
import datetime
import random

#from gross_stock_capture import HIDAR_EXCEL


def get_stock_urls(Stock_Num):

    urls = []
    #urls.append(f'http://kgieworld.moneydj.com/ZXQW/zc/zca/zca_{Stock_Num}.djhtm')
    #urls.append(f'https://kgieworld.moneydj.com/z/zc/zca/zca_{Stock_Num}.djhtm')
    #urls.append(f'http://jsjustweb.jihsun.com.tw/z/zc/zca/zca_{Stock_Num}.djhtm')
    urls.append(f'https://dj.mybank.com.tw/z/zc/zca/zca_{Stock_Num}.djhtm')
    #urls.append(f'https://tw.stock.yahoo.com/quote/{Stock_Num}') # Yahoo
    return urls

def get_reqs_data(urls):
    return [ requests.get(link) for link in urls ]

def parse_stock_price(reqs,dat_p,dat_v):

    for r in reqs:
        soup = BeautifulSoup(r.text,'html.parser')
        blocks = soup.find_all('table', {'class': 't01'})

        for blk in blocks:
            dat_p = float((((blk.find_all('tr')[1]).find_all('td')[7].text)).replace(',',''))
            dat_v = float((((blk.find_all('tr')[3]).find_all('td')[7].text)).replace(',',''))
    return dat_p,dat_v

def parse_yahoo_data(reqs,dat_p,dat_v):
    block0=['NA']
    for r in reqs:
        soup = BeautifulSoup(r.text,'html.parser')
        if soup.find_all('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c) C($c-trend-up)'}) is exists:
            block0 = soup.find_all('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c) C($c-trend-up)'})
        elif soup.find_all('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c) C($c-trend-down)'}) is exists:
            block0 = soup.find_all('span', {'class': 'Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c) C($c-trend-up)'})

        print(block0)

        block1 = soup.find_all('span', {'class': 'Fz(16px) C($c-link-text) Mb(4px)'})
        dat_p = float((block0[0].text).replace(',',''))
        dat_v = float((block1[0].text).replace(',',''))
    return dat_p,dat_v



# =====          =====
# =====   main   =====
# =====          =====
if __name__ == '__main__':

    HIDAR_EXCEL = 1
    RECAL_EXCEL = 1
    print('\n===================================')
    print(' Time : '+str(datetime.datetime.today()))
    print('===================================')
    start_time = time.time()

    path_xls = 'C:\\Users\\JS Wang\\Desktop\\test\\tmp.xlsx'
    path_fin = 'C:\\Users\JS Wang\Desktop\\test\gross_all_0115.txt'

    if( HIDAR_EXCEL == 1 ):

        fi_stk = open(path_fin,'r')
        lines = fi_stk.readlines()

        STK_CNT = 1
        STK_PRI = []
        STK_VOL = []

        reqs_date = requests.get("https://dj.mybank.com.tw/z/zc/zcl/zcl_2330.djhtm")
        soup_date = BeautifulSoup(reqs_date.text, 'html.parser')
        date_tmp = (((soup_date.find('table', {'class': 't01'})).find_all('tr')[7]).find_all('td')[0]).text
        STK_PRI.append(date_tmp)
        STK_VOL.append(date_tmp)

        js_tmp = [0,0]
        for line in lines:
            STK_NUM = int(line)
            dat_p = 0
            dat_v = 0
            js_tmp = parse_stock_price(get_reqs_data(get_stock_urls(str(STK_NUM))),dat_p,dat_v)
            #js_tmp = parse_yahoo_data(get_reqs_data(get_stock_urls(str(STK_NUM))),dat_p,dat_v)
            STK_PRI.append(float(js_tmp[0]))
            STK_VOL.append(float(js_tmp[1]))
            print(">> No."+str(STK_CNT) + "  ...  "+str(STK_NUM))

            #delay_choices = [0.1, 0.2, 0.3, 0.4 ,0.5]       # delay time
            #delay = random.choice(delay_choices)            # random choice
            #time.sleep(delay)                               # delay

            STK_CNT+=1

        print('\n\nWriting to the excel file ... '+str(path_xls)+'\n\n')
        wb = openpyxl.load_workbook(path_xls) if os.path.exists(path_xls) else openpyxl.Workbook()
        print(wb.sheetnames)

        st_flag=0
        for stn in wb.sheetnames: st_flag+=1 if stn == 'Price' else +0
        st0 = wb['Price'] if st_flag == 1 else wb.create_sheet('Price',0)
        st_flag=0
        for stn in wb.sheetnames: st_flag+=1 if stn == 'Volume' else +0
        st1 = wb['Volume'] if st_flag == 1 else wb.create_sheet('Volume',1)
        st_flag=0
        for stn in wb.sheetnames: st_flag+=1 if stn == 'Value' else +0
        st2 = wb['Value'] if st_flag == 1 else wb.create_sheet('Value',2)

        cnt_r0 = st0.max_row
        cnt_c0 = st0.max_column
        cnt_r1 = st1.max_row
        cnt_c1 = st1.max_column
        cnt_r2 = st2.max_row
        cnt_c2 = st2.max_column

        for r1 in range(1,len(STK_PRI)+1): (st0.cell(row=r1, column=cnt_c0+1)).value = STK_PRI[r1-1]
        for r2 in range(1,len(STK_VOL)+1): (st1.cell(row=r2, column=cnt_c1+1)).value = STK_VOL[r2-1]
        for r3 in range(1,len(STK_PRI)+1): (st2.cell(row=r3, column=cnt_c2+1)).value = STK_PRI[r3-1] if r3 == 1 else round((float(STK_PRI[r3-1])*float(STK_VOL[r3-1])/100000),2)
        print('\nFinished !!\n')
        wb.save(path_xls)
        #END of writing to the excel

    if( RECAL_EXCEL == 1 ):

        wb = openpyxl.load_workbook(path_xls)
        st0 = wb['Price']
        st1 = wb['Volume']
        st2 = wb['Value']
        cnt_r0 = st0.max_row
        cnt_c0 = st0.max_column
        cnt_r1 = st1.max_row
        cnt_c1 = st1.max_column
        cnt_r2 = st2.max_row
        cnt_c2 = st2.max_column

        #for ra in range(2,st0.max_row+1):
        #    for ca in range(1,st0.max_column+1):
        #        print(str((st0.cell(row=ra, column=ca)).value)+', ', end='')
        cal_js=1
        for val in range(2,st2.max_row+1):
            if (st2.cell(row=val, column=st2.max_column)).value > 1: # Value
                sum_lst=[0,0,0,0]
                day_lst=[3,5,10,20]
                avg_lst=[0,0,0,0,0]
                today_price=0
                for cal_r in range(val,val+1):
                    for cnt_d in range(len(day_lst)):
                        for cal_c in range(st0.max_column,st0.max_column-day_lst[cnt_d],-1):
                            sum_lst[cnt_d]+=(st0.cell(row=cal_r, column=cal_c)).value
                            if ((cal_c == st0.max_column) and (cnt_d == 0)):
                                today_price=sum_lst[cnt_d]


                    for cnt_d in range(4):
                        avg_lst[cnt_d] = sum_lst[cnt_d]/day_lst[cnt_d]
                        avg_lst[4] += avg_lst[cnt_d]
                        #print('>>> '+str(day_lst[cnt_d])+'-Day:'+str(avg_lst[cnt_d]), end='')
                    avg_lst[4] /= 4

                    cal_tmp=0
                    for cnt_d in range(4):
                        if ((today_price/avg_lst[cnt_d] > 0.9) and (today_price/avg_lst[cnt_d] < 1.1)): cal_tmp+=1 # diff ratio
                    if cal_tmp == 4:
                        if ( (avg_lst[0] > avg_lst[1]) and (avg_lst[1] > avg_lst[2]) and (avg_lst[2] > avg_lst[3]) and (today_price < 500) ):
                            print('>> ', end='')
                            print('%8s' % str((st0.cell(row=cal_r, column=2)).value), end='')
                            print('  Today :%6s' % str(today_price), end='')
                            print('%8s' % str(round(avg_lst[0],2))+'元'+'(%2s' % str(day_lst[0])+'日)', end='')
                            print('%8s' % str(round(avg_lst[1],2))+'元'+'(%2s' % str(day_lst[1])+'日)', end='')
                            print('%8s' % str(round(avg_lst[2],2))+'元'+'(%2s' % str(day_lst[2])+'日)', end='')
                            print('%8s' % str(round(avg_lst[3],2))+'元'+'(%2s' % str(day_lst[3])+'日)', end='')
                            print('%10s' % str((st2.cell(row=val, column=st2.max_column)).value)+str('億'), end='')
                            print('%10s' % str((st0.cell(row=cal_r, column=3)).value), end='')
                            print()
            cal_js+=1
        wb.save(path_xls)


    end_time = time.time()
    print('\nTake time : '+str(round((end_time-start_time),2))+'(S)')

    print()
    print("       *******             *****        *             *     ***********                         *      ")
    print("     *         *         *       *      * *           *    *                                   *       ")
    print("     *          *       *         *     *  *          *    *                                  *        ")
    print("     *           *     *           *    *   *         *    *                                 *         ")
    print("     *            *    *           *    *    *        *    *                                *          ")
    print("     *            *    *           *    *     *       *    *                               *           ")
    print("     *            *    *           *    *      *      *     ***********      *            *            ")
    print("     *            *    *           *    *       *     *    *                  *          *             ")
    print("     *            *    *           *    *        *    *    *                   *        *              ")
    print("     *           *     *           *    *         *   *    *                    *      *               ")
    print("     *          *       *         *     *          *  *    *                     *    *                ")
    print("     *         *         *       *      *           * *    *                      *  *                 ")
    print("       *******             *****        *             *     ***********            *                   ")
    print()
