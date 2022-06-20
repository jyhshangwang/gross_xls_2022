import gross_sub as sub


def get_yahoo_urls(Stock_Num):
    urls = []
    urls.append(f'https://tw.stock.yahoo.com/quote/{Stock_Num}')
    return urls

def get_yahoo_urls_asynch(lines):
    urls = []
    for line in lines:
        Stock_Num = str(int(line))
        urls.append(f'https://tw.stock.yahoo.com/quote/{Stock_Num}')
    return urls

def yahoo_stock_data():

    fi_stk = open('C:\\Users\\JS Wang\\Desktop\\test\\gross_all_chk.txt','r')
    fo_stk = open('C:\\Users\\JS Wang\\Desktop\\test\\test_mode_out.txt','w')
    lines = fi_stk.readlines()
    for line in lines:
        STK_NUM = int(line)
        print(str(STK_NUM)+' .. ', end='')
        y_dat = sub.parse_stock_data_yahoo(sub.get_reqs_data(get_yahoo_urls(str(STK_NUM))))
        print(y_dat)
        print(y_dat, file=fo_stk)
    fi_stk.close()
    fo_stk.close()


def yahoo_stock_data_asynch():

    fi_stk = open('C:\\Users\\JS Wang\\Desktop\\test\\gross_all_chk.txt','r')
    fo_stk = open('C:\\Users\\JS Wang\\Desktop\\test\\test_mode_out.txt','w')
    lines = fi_stk.readlines()
    idlst = []
    for line in lines: idlst.append(int(line))
    YAHOO_LST = [0,0]
    YAHOO_LST = sub.parse_yahoo_asynch(sub.get_reqs_data_asynch(get_yahoo_urls_asynch(lines)),len(idlst))
    YAHOO_PRI = []
    for cnt_i in range(0,len(idlst)):
        tmpid = idlst[cnt_i]
        for cnt_j in range(0,len(YAHOO_LST[0])):
            if YAHOO_LST[0][cnt_j] == tmpid:
                print('\r'+str(YAHOO_LST[0][cnt_j]), end='')
                YAHOO_PRI.append(str((YAHOO_LST[1][cnt_j]).replace(',','')))
                break
    print()
    for i in range(len(YAHOO_PRI)): print(YAHOO_PRI[i], file=fo_stk)
    
    fi_stk.close()
    fo_stk.close()
