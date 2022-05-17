import gross_sub as sub


def get_yahoo_urls(Stock_Num):
    urls = []
    urls.append(f'https://tw.stock.yahoo.com/quote/{Stock_Num}') # Yahoo
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

