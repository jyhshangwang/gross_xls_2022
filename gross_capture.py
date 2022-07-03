from bs4 import BeautifulSoup
from openpyxl.workbook.workbook import Workbook
import grequests
import requests
import pyquery
import time
import datetime
from fake_useragent import UserAgent
import random
import openpyxl
import winsound
import gross_class as cls
from alive_progress import alive_bar
import gross_sub as sub
import loguru


class Dayilyinfo:
    def __init__(self, \
                op_price,hi_price,lo_price,td_price, \
                up_dn_price,mx_price,mi_price, \
                pe_ratio,mx_volum,mi_volum,td_volum, \
                dividend, \
                inc_year,inc_5day,inc_1mon,inc_2mon,inc_3mon, \
                roe_rate,stock_cnt,ratio_cmt \
        ):
        self.Op_price = float(op_price.replace(',',''))
        self.Hi_price = float(hi_price.replace(',',''))
        self.Lo_price = float(lo_price.replace(',',''))
        self.Td_price = float(td_price.replace(',',''))
        self.Up_dn_price = float(up_dn_price)
        self.Mx_price = float(mx_price.replace(',',''))
        self.Mi_price = float(mi_price.replace(',',''))
        self.Pe_ratio = float(pe_ratio.replace(',','')) if pe_ratio != 'N/A' else 'NEG'
        self.Mx_volum = int(mx_volum.replace(',',''))
        self.Mi_colum = int(mi_volum.replace(',',''))
        self.Td_volum = int(td_volum.replace(',',''))
        self.Dividend = dividend
        self.Inc_year = inc_year
        self.Inc_5day = inc_5day
        self.Inc_1mon = inc_1mon
        self.Inc_2mon = inc_2mon
        self.Inc_3mon = inc_3mon
        self.Roe_rate = roe_rate
        self.Stock_cnt = float(stock_cnt.replace(',',''))
        self.Ratio_cmt = ratio_cmt.replace('、',' ')
    def __repr__(self):
        return [self.Ratio_cmt,self.Pe_ratio,self.Mx_volum,self.Mi_colum,self.Td_volum,self.Stock_cnt,self.Up_dn_price,\
                self.Mx_price,self.Mi_price,self.Td_price,\
                self.Inc_5day,self.Inc_1mon,self.Inc_2mon,self.Inc_3mon,self.Inc_year\
            ]
    def turnover(self): return round(self.Td_volum/(self.Stock_cnt*100),2)
    def yd_price(self): return self.Td_price-self.Up_dn_price
    def is_jump(self): return self.Op_price-self.yd_price()
    def red_k(self): return self.Td_price-self.Op_price
    def diff_rate(self): return round((self.Hi_price-self.Lo_price)/self.yd_price(),2)
    def inc_1day(self): return str(round(self.Up_dn_price/self.yd_price()*100,2))+'%'
    def val(self): return round(self.Td_price*self.Td_volum/100000,2)


def file_open_init():
    path_finn  = 'C:\\Users\\JS Wang\\Desktop\\test\\gross_all_0115.txt'
    path_fout1 = 'C:\\Users\\JS Wang\\Desktop\\data\\__stock__data_d1.txt'
    path_fout2 = 'C:\\Users\\JS Wang\\Desktop\\data\\__stock__data_d2.txt'
    path_fout3 = 'C:\\Users\\JS Wang\\Desktop\\data\\__stock__data_d3.txt'

    fin  = open(path_finn, 'r')
    fout1 = open(path_fout1, 'w', encoding='UTF-8')
    fout2 = open(path_fout2, 'w', encoding='UTF-8')
    fout3 = open(path_fout3, 'w', encoding='UTF-8')

    return [fin,fout1,fout2,fout3]


def file_close(lst):
    for i in range(len(lst)): lst[i].close()


def beep():
    duration = 2000 # mS
    freq = 400      # Hz
    for i in range(1):
        if i%2 == 0 : winsound.Beep(freq, duration)
    time.sleep(0.2)


def get_urls_part3(Stock_Num):
    urls_P3=[]
    urls_P3.append(f'http://jsjustweb.jihsun.com.tw/z/zc/zcn/zcn_{Stock_Num}.djhtm')
    #urls_P3.append(f'https://dj.mybank.com.tw/z/zc/zcn/zcn_{Stock_Num}.djhtm')
    return urls_P3


def get_urls_lst(urls,num):
    if num == 0: urls_lst = [ f'https://kgieworld.moneydj.com/z/zc/zcl/zcl_{url}.djhtm' for url in urls ]
    if num == 1: urls_lst = [ f'http://jsjustweb.jihsun.com.tw/z/zc/zce/zce_{url}.djhtm' for url in urls ]
    if num == 2: urls_lst = [ f'http://jsjustweb.jihsun.com.tw/z/zc/zca/zca_{url}.djhtm' for url in urls ]
    if num == 3: urls_lst = [ f'https://kgieworld.moneydj.com/z/zc/zcl/zcl_{url}.djhtm' for url in urls ]
    return urls_lst


def tpinv(str1):
    return float((str1.replace('%','')).replace(',','')) if str1 != '' else 0


def get_reqs_data_asynch(urls):
    reqs = ( grequests.get(url) for url in urls )
    response = grequests.imap(reqs, grequests.Pool(len(urls)))
    return response


def get_reqs_data(urls):
    reqs = [ requests.get(url) for url in urls ]
    return reqs


@loguru.logger.catch
def parse_data_1(response,cnt_stk): # Season
    season_lst = []
    bar = cls.ProgressBar(cnt_stk)
    for r in response:
        d = pyquery.PyQuery(r.text)
        tits = (list(d('title').items()))[0].text().strip().replace('個股獲利能力-','')
        tbls = list(d('table').items())
        tbls = tbls[2:3]
        for tbl in tbls:
            trs = list(tbl('tr').items())
            trs = trs[3:7]
            sea_lst = []
            for tr in trs:
                tds = list(tr('td').items())
                sea_lst.append( [ tds[num].text().strip() for num in range(len(tds)) if num == 0 or num == 4 or num == 6 or num == 10 ] )
            for i in range(len(sea_lst[0:1])):
                gross_diff = str(round(tpinv(sea_lst[i][1])-tpinv(sea_lst[i+1][1]), 2))+'%'
                profit_diff = str(round(tpinv(sea_lst[i][2])-tpinv(sea_lst[i+1][2]), 2))+'%'
                eps_rate = round(( float(sea_lst[i][-1]) - float(sea_lst[i+1][-1]) ) / float(sea_lst[i+1][-1]), 2) if float(sea_lst[i+1][-1]) > 0 else 'NEG'
            eps_sum = 0
            for i in range(len(sea_lst)): eps_sum += float(sea_lst[i][-1])
        dat_lst = []
        dat_lst.extend(sea_lst[0])
        dat_lst.insert(2,gross_diff)
        dat_lst.insert(4,profit_diff)
        dat_lst.append(sea_lst[1][-1])
        dat_lst.append(sea_lst[2][-1])
        dat_lst.append(sea_lst[3][-1])
        dat_lst.append(str(eps_rate))
        dat_lst.append(str(eps_sum))
        str_dat = ';'.join(dat_lst)
        season_lst.append([int(tits),str_dat])
        bar.update()
    return season_lst


@loguru.logger.catch
def parse_data_2(response,cnt_stk): # Daily
    daily_lst = []
    bar = cls.ProgressBar(cnt_stk)
    for r in response:
        d = pyquery.PyQuery(r.text)
        tits = (list(d('title').items()))[0].text().strip().replace('個股基本資料-','')
        tbls = list(d('table').items())
        tbls = tbls[2:3]
        for tbl in tbls:
            trs = list(tbl('tr').items())
            trs = trs[1:24]
            trs = trs[0:3]+trs[4:5]+trs[6:11]+trs[12:13]+trs[19:20]
            for tr in trs:
                if tr == trs[0]: # price
                    tds = list(tr('td').items())
                    op_price = tds[1].text().strip()
                    hi_price = tds[3].text().strip()
                    lo_price = tds[5].text().strip()
                    td_price = tds[7].text().strip()
                if tr == trs[1]:
                    tds = list(tr('td').items())
                    up_dn_price = tds[1].text().strip()
                    mx_price = tds[3].text().strip()
                    mi_price = tds[5].text().strip()
                if tr == trs[2]:
                    tds = list(tr('td').items())
                    pe_ratio = tds[1].text().strip()
                    mx_volum = tds[3].text().strip()
                    mi_volum = tds[5].text().strip()
                    td_volum = tds[7].text().strip()
                if tr == trs[3]:
                    tds = list(tr('td').items())
                    dividend = tds[1].text().strip()
                if tr == trs[4]:
                    tds = list(tr('td').items())
                    inc_year = tds[1].text().strip()
                if tr == trs[5]:
                    tds = list(tr('td').items())
                    inc_5day = tds[1].text().strip()
                if tr == trs[6]:
                    tds = list(tr('td').items())
                    inc_1mon = tds[1].text().strip()
                if tr == trs[7]:
                    tds = list(tr('td').items())
                    inc_2mon = tds[1].text().strip()
                if tr == trs[8]:
                    tds = list(tr('td').items())
                    inc_3mon = tds[1].text().strip()
                    roe_rate = tds[5].text().strip()
                if tr == trs[9]:
                    tds = list(tr('td').items())
                    stock_cnt = tds[1].text().strip()
                if tr == trs[10]:
                    tds = list(tr('td').items())
                    ratio_cmt = tds[1].text().strip()
            stk = Dayilyinfo(op_price,hi_price,lo_price,td_price,up_dn_price,mx_price,mi_price,\
                                pe_ratio,mx_volum,mi_volum,td_volum,dividend,\
                                inc_year,inc_5day,inc_1mon,inc_2mon,inc_3mon,roe_rate,stock_cnt,ratio_cmt)
        dat_lst = []
        dat_lst = stk.__repr__()
        dat_lst.insert(5,'')
        dat_lst.insert(6,'')
        dat_lst.insert(7,stk.turnover())
        dat_lst.insert(12,stk.is_jump())
        dat_lst.insert(13,stk.red_k())
        dat_lst.insert(14,stk.diff_rate())
        dat_lst.insert(16,stk.yd_price())
        dat_lst.insert(17,stk.inc_1day())
        dat_lst.append(stk.val())
        for i in range(len(dat_lst)): dat_lst[i] = str(dat_lst[i])
        str_dat = ';'.join(dat_lst)
        daily_lst.append([int(tits),str_dat])
        bar.update()
    return daily_lst


@loguru.logger.catch
def parse_data_3(response,stk_cnt):
    tmp_lst = []
    bar = cls.ProgressBar(stk_cnt)
    for r in response:
        d = pyquery.PyQuery(r.text)
        tits = (list(d('title').items()))[0].text().strip().replace('個股法人持股-','')
        tbls = list(d('table').items())
        tbls = tbls[2:3]
        counter_lst = []
        for tbl in tbls:
            trs = list(tbl('tr').items())
            trs = trs[7:13]
            for tr in trs:
                tds = list(tr('td').items())
                if tds[1].text().strip() != '--':
                    cnt_lst = [ int((tds[i].text().strip()).replace(',','')) for i in range(len(tds)) if 0 < i < 5 ]
                else:
                    cnt_lst = [ 0 for i in range(len(tds)) if 0 < i < 5 ]
                counter_lst.append(cnt_lst)
        cmt = ''
        for i in range(len(counter_lst)-1):
            if   counter_lst[i][1] >  0 : cmt+='+'
            elif counter_lst[i][1] == 0 : cmt+='.'
            else                        : cmt+='-'
        if   cmt[0:5] == '+++++'                 : cmt = '連5買'
        elif cmt[0:4] == '++++' and cmt[4] != '+': cmt = '連4買'
        elif cmt[0:3] == '+++'  and cmt[3] != '+': cmt = '連3買'
        elif cmt[0:2] == '++'   and cmt[2] != '+': cmt = '連2買'
        elif cmt[0:1] == '+'    and cmt[1] != '+': cmt = '連1買'
        elif cmt[0:1] == '-'    and cmt[1] != '-': cmt = '賣1日'
        elif cmt[0:2] == '--'   and cmt[2] != '-': cmt = '賣2日'
        elif cmt[0:3] == '---'  and cmt[3] != '-': cmt = '賣3日'
        elif cmt[0:4] == '----' and cmt[4] != '-': cmt = '賣4日'
        elif cmt[0:5] == '-----'                 : cmt = '賣5日'
        elif cmt[0:5] == '.....'                 : cmt = 'NA'
        else                                     : cmt = 'No rule'

        for i in range(len(counter_lst)):
            for j in range(len(counter_lst[i])):
                counter_lst[i][j] = str(counter_lst[i][j])

        counter_lst = [ ';'.join(counter_lst[i]) for i in range(len(counter_lst)) ]
        str_dat = ';'.join(counter_lst)
        str_dat = str_dat+';'+cmt
        tmp_lst.append([int(tits),str_dat])
        bar.update()
    return tmp_lst


def parse_part0_M(reqs):

    STK_REV = []
    yoy_rate  = [0,0,0,0]
    m_ratio   = [1,1,1,10/8,1,1,1]
    month_rev = [0,0,0,0,0,0,0]
    rev_tmp   = [0,0,0,0,0,0,0]
    #reqs_revenue  = requests.get(f'http://kgieworld.moneydj.com/ZXQW/zc/zch/zch_{Stock_Num}.djhtm')
    reqs_revenue  = requests.get(f'https://dj.mybank.com.tw/z/zc/zch/zch_{Stock_Num}.djhtm')
    #reqs_revenue  = requests.get(f'http://jsjustweb.jihsun.com.tw/z/zc/zch/zch_{Stock_Num}.djhtm')
    soup_revenue = BeautifulSoup(reqs_revenue.text,'html.parser')

    for m1 in range(0,len(month_rev)): month_rev[m1] = 0 if( (((soup_revenue.find_all("table")[1]).find_all("tr")[m1+7]).find_all("td")[1].text).replace(',','') == '' ) else int((((soup_revenue.find_all("table")[1]).find_all("tr")[m1+7]).find_all("td")[1].text).replace(',',''))
    for y1 in range(0,len(yoy_rate)):  yoy_rate[y1] = 0 if( (((((soup_revenue.find_all("table")[1]).find_all("tr")[y1+7]).find_all("td")[4]).text).replace('%','')).replace(',','') == '' ) else float((((((soup_revenue.find_all("table")[1]).find_all("tr")[y1+7]).find_all("td")[4]).text).replace('%','')).replace(',',''))

    for cnt_rev in range(0,len(month_rev)): rev_tmp[cnt_rev] = month_rev[cnt_rev]*float(m_ratio[cnt_rev])

    if( (month_rev[3]+month_rev[4]+month_rev[5]) == 0 or (month_rev[4]+month_rev[5]+month_rev[6]) == 0 ):
        STK_REV.append('NA')
        STK_REV.append('NA')
    else:
        rev_dvo1 = ( round(( ( (rev_tmp[0]+rev_tmp[1]+rev_tmp[2]) - (rev_tmp[3]+rev_tmp[4]+rev_tmp[5]) ) / (rev_tmp[3]+rev_tmp[4]+rev_tmp[5]) ), 2) )*100
        rev_dvo2 = ( round(( ( (rev_tmp[1]+rev_tmp[2]+rev_tmp[3]) - (rev_tmp[4]+rev_tmp[5]+rev_tmp[6]) ) / (rev_tmp[4]+rev_tmp[5]+rev_tmp[6]) ), 2) )*100
        rev_dvrms = round((rev_dvo1 - rev_dvo2),2)
        STK_REV.append(str(rev_dvo1)+'%')
        STK_REV.append(str(rev_dvrms)+'%')

    if  ( rev_tmp[0] > rev_tmp[1] and rev_tmp[1] > rev_tmp[2] and rev_tmp[2] > rev_tmp[3] ): STK_REV.append('營收連3增')
    elif( rev_tmp[0] > rev_tmp[1] and rev_tmp[1] > rev_tmp[2] and rev_tmp[2] < rev_tmp[3] ): STK_REV.append('營收連2增')
    elif( rev_tmp[0] > rev_tmp[1] and rev_tmp[1] < rev_tmp[2] and rev_tmp[2] < rev_tmp[3] ): STK_REV.append('營收月增1')
    elif( rev_tmp[0] < rev_tmp[1] and rev_tmp[1] > rev_tmp[2] and rev_tmp[2] > rev_tmp[3] ): STK_REV.append('營收月-1')
    elif( rev_tmp[0] < rev_tmp[1] and rev_tmp[1] < rev_tmp[2] and rev_tmp[2] > rev_tmp[3] ): STK_REV.append('營收連-2')
    elif( rev_tmp[0] < rev_tmp[1] and rev_tmp[1] < rev_tmp[2] and rev_tmp[2] < rev_tmp[3] ): STK_REV.append('營收連-3')
    else:                                                                                    STK_REV.append('NA')

    STK_REV.append(((soup_revenue.find_all("table")[1]).find_all("tr")[7]).find_all("td")[0].text) # Date
    STK_REV.append(str(round(float(month_rev[0]/100000),2)))                                       # Revenue for this month
    STK_REV.append(((soup_revenue.find_all("table")[1]).find_all("tr")[7]).find_all("td")[2].text) # MOM
    STK_REV.append(((soup_revenue.find_all("table")[1]).find_all("tr")[7]).find_all("td")[4].text) # YOY
    STK_REV.append(((soup_revenue.find_all("table")[1]).find_all("tr")[7]).find_all("td")[6].text) # Total YOY

    y_cnt=0
    for y in range(0,len(yoy_rate)-1):
        if( yoy_rate[y] < yoy_rate[y+1] ):
            tmp = yoy_rate[y]
            yoy_rate[y] = yoy_rate[y+1]
            yoy_rate[y+1] = tmp
            y_cnt+=1
    if  ( y_cnt == 3               ): STK_REV.append(str( round( int(yoy_rate[3]) , -1 ) )+'%'+'往下')
    elif( y_cnt == 2 or y_cnt == 1 ): STK_REV.append(str( round( int(yoy_rate[3]) , -1 ) )+'%'+'持平')
    elif( y_cnt == 0               ): STK_REV.append(str( round( int(yoy_rate[3]) , -1 ) )+'%'+'往上')

    return STK_REV


def parse_part3(reqs):

    for r in reqs:
        soup = BeautifulSoup(r.text, 'html.parser')
        blocks = soup.find_all('table', {'class' :'t01'})

        FIN_LIST_05 = []
        FIN_LIST_07 = []
        FIN_LIST_12 = []
        FIN_LIST_13 = []
        FIN_SUM = [0,1,2,3,4,5]
        for block in blocks:
            for fin_i in range(7,12,1): # 5 times
                if( len(block.find_all('tr')) != 13 or ((((block.find_all('tr')[fin_i]).find_all('td')[ 7]).text) == '') or ((((block.find_all('tr')[fin_i]).find_all('td')[13]).text) == '') ):
                    FIN_LIST_05.append('0')
                    FIN_LIST_07.append('0')
                    FIN_LIST_12.append('0')
                    FIN_LIST_13.append('0')
                else:
                    FIN_LIST_05.append((((block.find_all('tr')[fin_i]).find_all('td')[ 5]).text).replace(',','')) # finacial value
                    FIN_LIST_07.append((((block.find_all('tr')[fin_i]).find_all('td')[ 7]).text).replace(',','')) # finacial rate
                    FIN_LIST_12.append((((block.find_all('tr')[fin_i]).find_all('td')[12]).text).replace(',','')) # short selling value
                    FIN_LIST_13.append((((block.find_all('tr')[fin_i]).find_all('td')[13]).text).replace(',','')) # short selling rate

            FIN_SUM[0] = int(FIN_LIST_05[0])
            FIN_SUM[1] = int(FIN_LIST_05[0])+int(FIN_LIST_05[1])+int(FIN_LIST_05[2])+int(FIN_LIST_05[3])+int(FIN_LIST_05[4])
            FIN_SUM[2] = round((float(FIN_LIST_07[0].replace('%',''))-float(FIN_LIST_07[4].replace('%',''))),2)
            FIN_SUM[3] = int(FIN_LIST_12[0])
            FIN_SUM[4] = round((float(FIN_LIST_13[0].replace('%',''))-float(FIN_LIST_13[4].replace('%',''))),2)

            if( FIN_SUM[0] < 0 and FIN_SUM[1] < 0 and FIN_SUM[2] < 0 and FIN_SUM[3] > 0 and FIN_SUM[4] > 0 ):
                FIN_SUM[5] = '資減券增'
            elif ( FIN_SUM[0] > 0 and FIN_SUM[1] > 0 and (FIN_SUM[0]/FIN_SUM[1]) > 0.7 ):
                FIN_SUM[5] = '資暴增!'
            else:
                FIN_SUM[5] = 'NA'
    return FIN_SUM


def main():

    sub.get_stock_datetime()
    start_time = time.time()

    file = file_open_init()
    stocks = file[0].readlines()
    stock_num_lst = [ int(stock) for stock in stocks ]

    STK_DAT1 = parse_data_1(get_reqs_data_asynch(get_urls_lst(stock_num_lst,1)),len(stock_num_lst))
    for stock in stocks:
        for i in range(len(STK_DAT1)):
            if int(stock) == STK_DAT1[i][0]:
                #print('\r'+str(stock), end='')
                print(str(STK_DAT1[i][0])+';'+STK_DAT1[i][1], file=file[1])
                break

    STK_DAT2 = parse_data_2(get_reqs_data_asynch(get_urls_lst(stock_num_lst,2)),len(stock_num_lst))
    for stock in stocks:
        for i in range(len(STK_DAT2)):
            if int(stock) == STK_DAT2[i][0]:
                #print('\r'+str(stock), end='')
                print(str(STK_DAT2[i][0])+';'+STK_DAT2[i][1], file=file[2])
                break

    STK_DAT3 = parse_data_3(get_reqs_data_asynch(get_urls_lst(stock_num_lst,3)),len(stock_num_lst))
    for stock in stocks:
        for i in range(len(STK_DAT3)):
            if int(stock) == STK_DAT3[i][0]:
                #print('\r'+str(stock), end='')
                print(str(STK_DAT3[i][0])+';'+STK_DAT3[i][1], file=file[3])
                break

    file_close(file)
    '''
    fi1 = open("C:\\Users\\JS Wang\\Desktop\\data\\__stock__data_d1.txt",'r',encoding='utf-8')
    fi2 = open("C:\\Users\\JS Wang\\Desktop\\data\\__stock__data_d2.txt",'r',encoding='utf-8')
    fi3 = open("C:\\Users\\JS Wang\\Desktop\\data\\__stock__data_d3.txt",'r',encoding='utf-8')
    fot = open("C:\\Users\\JS Wang\\Desktop\\data\\__stock__data_sum.txt",'w',encoding='utf-8')

    buf1 = fi1.read()
    buf2 = fi2.read()
    buf3 = fi3.read()
    sum = buf1.replace('\n','')+';'+buf2.replace('\n','')+';'+buf3.replace('\n','')
    fot.write(sum)

    fi1.close()
    fi2.close()
    fi3.close()
    fot.close()
    '''
    end_time = time.time()

    print('>> 運算時間 : '+f'{round(end_time-start_time,2)}'+'(S) => '+ \
                            f'{round((end_time-start_time)/60,2)}'+'(min).')



if __name__ == '__main__':

    with alive_bar(4, title='>> Generating...', length=40, bar='bubbles', spinner='radioactive') as bar:
        for i in range(4):
            time.sleep(.5)
            bar()
    print()

    loguru.logger.add(
        f'_daily_reccord_dlg{datetime.date.today():%Y%m%d}.log',
        rotation='1 day',
        retention='7 days',
        level='DEBUG'
    )

    main()

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

    beep()