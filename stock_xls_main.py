import loguru
import time
import datetime
import pandas as pd
import gross_sub as sub
import gross_stkshm as shm
import gross_yahoo as yas
import gross_class as cls
from openpyxl.styles import Font
import winsound
from alive_progress import alive_bar


# ======        ======
# ======  main  ======
# ======        ======
if __name__ == '__main__':

    with alive_bar(50, title='>> Starting stock program', length=40, bar='blocks') as bar:
        for i in range(50):
            time.sleep(.05)
            bar()

    para = input('\n>> Enter parameter : ')
    para1 = int(para[0]) #input('>> Capture data   1(YES)/0(NO) : ')
    para2 = int(para[1]) #input('>> Calculate data 1(YES)/0(NO) : ')
    para3 = int(para[2]) #input('>> Yahoo data     1(YES)/0(NO) : ')
    #loguru.logger.add( f'Stock_datalog_{datetime.date.today():%Y%m%d}.log', rotation='1 day', retention='7 days', level='DEBUG')
    loguru.logger.add(f'Stock_info_datalog.log', rotation='1 day', retention='7 days', level='DEBUG')
    YSTK_M = int(para3)
    TEST_M = 0
    DEMO_M = 0
    ASYN_M = 1
    HIDAR_EXCEL = [int(para1),int(para2)]
    sub.time_title()
    start_time = time.time()
    date_tmp = sub.get_stock_datetime()

    path_xls = 'C:\\Users\\JS Wang\\Desktop\\test\\tmp.xlsx'
    path_xls_tdout = 'C:\\Users\\JS Wang\\Desktop\\test\\output.xlsx'
    path_fin = 'C:\\Users\\JS Wang\\Desktop\\test\\gross_all_0115.txt'

    if( TEST_M == 1 ): sub.revenue_info('https://dj.mybank.com.tw/z/zc/zch/zch_3006.djhtm')
    if( YSTK_M == 1 ): yas.yahoo_stock_data()
    if( DEMO_M == 1 ): shm.display(path_xls)
    if( HIDAR_EXCEL[0] == 1 ):

        fi_stk = open(path_fin,'r')
        lines = fi_stk.readlines()

        STK_CNT = 1
        STK_PRI = []
        STK_VOL = []
        STK_TRR = []
        STK_PRI.append(date_tmp)
        STK_VOL.append(date_tmp)
        STK_TRR.append(date_tmp)


        if( ASYN_M == 0 ):
            JS_TMP = [0,0,0]
            for line in lines:
                STK_NUM = int(line)
                JS_TMP = sub.parse_stock_data(sub.get_reqs_data(sub.get_stock_urls(str(STK_NUM))))
                STK_PRI.append(float(JS_TMP[0]))
                STK_VOL.append(float(JS_TMP[1]))
                STK_TRR.append(float(JS_TMP[2]))
                print(">> No."+str(STK_CNT) + "  ...  "+str(STK_NUM))
                #sub.rand_on()
                STK_CNT+=1

        if( ASYN_M == 1 ):
            idlst = []
            links = []
            JS_TMP2 = [0,0,0,0]
            for line in lines: idlst.append(int(line))
            for line in lines: links.append(f'https://dj.mybank.com.tw/z/zc/zca/zca_{str(int(line))}.djhtm')
            JS_TMP2 = sub.parse_stock_data_asynch(sub.get_reqs_data_asynch(links),len(idlst))
            loguru.logger.info('Sort the data now ...')
            for cnt_i in range(0,len(idlst)):
                tmpid = idlst[cnt_i]
                for cnt_j in range(0,len(JS_TMP2[0])):
                    if JS_TMP2[0][cnt_j] == tmpid:
                        print('\r'+str(JS_TMP2[0][cnt_j]), end='')
                        STK_PRI.append(float(JS_TMP2[1][cnt_j]))
                        STK_VOL.append(float(JS_TMP2[2][cnt_j]))
                        STK_TRR.append(float(JS_TMP2[3][cnt_j]))
                        break
            print()


        loguru.logger.info('>> STEP1. Write to  ... '+str(path_xls))
        wb = sub.xls_wb_on(path_xls)
        st0 = sub.xls_st_on(wb,0,'Price'   ,0)
        st1 = sub.xls_st_on(wb,0,'Volume'  ,1)
        st2 = sub.xls_st_on(wb,0,'Value'   ,2)
        st3 = sub.xls_st_on(wb,0,'Turnover',3)

        cnt_c0 = st0.max_column
        cnt_c1 = st1.max_column
        cnt_c2 = st2.max_column
        cnt_c3 = st3.max_column

        for r1 in range(1,len(STK_PRI)+1): (st0.cell(row=r1, column=cnt_c0+1)).value = STK_PRI[r1-1]
        for r2 in range(1,len(STK_VOL)+1): (st1.cell(row=r2, column=cnt_c1+1)).value = STK_VOL[r2-1]
        for r3 in range(1,len(STK_PRI)+1): (st2.cell(row=r3, column=cnt_c2+1)).value = STK_PRI[r3-1] if r3 == 1 else round((float(STK_PRI[r3-1])*float(STK_VOL[r3-1])/100000),2)
        for r4 in range(1,len(STK_TRR)+1): (st3.cell(row=r4, column=cnt_c3+1)).value = STK_TRR[r4-1]

        loguru.logger.success('Completion OK: Capture daily info.')
        sub.xls_wb_off(wb,path_xls)


    if( HIDAR_EXCEL[1] == 1 ):


        loguru.logger.info('>> STEP2. Calculate average line price from  ... '+str(path_xls))
        wb = sub.xls_wb_on(path_xls)
        st0 = sub.xls_st_on(wb,0,'Price',0)
        st4 = sub.xls_st_on(wb,0,'3ma'  ,4)
        st5 = sub.xls_st_on(wb,0,'5ma'  ,5)
        st6 = sub.xls_st_on(wb,0,'10ma' ,6)
        st7 = sub.xls_st_on(wb,0,'20ma' ,7)
        st8 = sub.xls_st_on(wb,0,'40ma' ,8)
        st9 = sub.xls_st_on(wb,0,'60ma' ,9)
        stt = sub.xls_st_on(wb,0,'Tangled',20)
        cnt_c0 = st0.max_column
        cnt_c4 = st4.max_column
        cnt_c5 = st5.max_column
        cnt_c6 = st6.max_column
        cnt_c7 = st7.max_column
        cnt_c8 = st8.max_column
        cnt_c9 = st9.max_column
        cnt_ct = stt.max_column
        pgs1 = cls.ProgressBar(st0.max_row)
        for r in range(1,st0.max_row+1):
            day_lst=[ 3, 5,10,20,40,60]
            avg_lst=[ 0, 0, 0, 0, 0, 0]
            if r != 1 :
                avg_lst=sub.cal_avg_price(st0,day_lst,r)
                avg_cmt=sub.cal_moving_average_tangled(avg_lst)
            pgs1.update()
            (st4.cell(row=r, column=cnt_c4+1)).value = avg_lst[0] if r != 1 else date_tmp
            (st5.cell(row=r, column=cnt_c5+1)).value = avg_lst[1] if r != 1 else date_tmp
            (st6.cell(row=r, column=cnt_c6+1)).value = avg_lst[2] if r != 1 else date_tmp
            (st7.cell(row=r, column=cnt_c7+1)).value = avg_lst[3] if r != 1 else date_tmp
            (st8.cell(row=r, column=cnt_c8+1)).value = avg_lst[4] if r != 1 else date_tmp
            (st9.cell(row=r, column=cnt_c9+1)).value = avg_lst[5] if r != 1 else date_tmp
            (stt.cell(row=r, column=cnt_ct+1)).value = avg_cmt    if r != 1 else date_tmp
        loguru.logger.success('Completion OK: Average line price')
        sub.xls_wb_off(wb,path_xls)


        loguru.logger.info('>> STEP3. Calculate Increase rate from  ... '+str(path_xls))
        wb = sub.xls_wb_on(path_xls)
        st0 = sub.xls_st_on(wb,0,'Price', 0)
        st1 = sub.xls_st_on(wb,0,'Inc1' ,10)
        st2 = sub.xls_st_on(wb,0,'Inc3' ,11)
        st3 = sub.xls_st_on(wb,0,'Inc5' ,12)
        st4 = sub.xls_st_on(wb,0,'Inc10',13)
        st5 = sub.xls_st_on(wb,0,'Inc20',14)
        st6 = sub.xls_st_on(wb,0,'Inc60',15)
        cnt_st1 = st1.max_column
        cnt_st2 = st2.max_column
        cnt_st3 = st3.max_column
        cnt_st4 = st4.max_column
        cnt_st5 = st5.max_column
        cnt_st6 = st6.max_column
        pgs2 = cls.ProgressBar(st0.max_row)
        for r in range(1,st0.max_row+1):
            day_lst=[ 1, 3, 5,10,20,60]
            rat_lst=[ 0, 0, 0, 0, 0, 0]
            if r!= 1 : rat_lst=sub.cal_increase_rate(st0,day_lst,r)
            pgs2.update()
            (st1.cell(row=r, column=cnt_st1+1)).value = rat_lst[0] if r != 1 else date_tmp
            (st2.cell(row=r, column=cnt_st2+1)).value = rat_lst[1] if r != 1 else date_tmp
            (st3.cell(row=r, column=cnt_st3+1)).value = rat_lst[2] if r != 1 else date_tmp
            (st4.cell(row=r, column=cnt_st4+1)).value = rat_lst[3] if r != 1 else date_tmp
            (st5.cell(row=r, column=cnt_st5+1)).value = rat_lst[4] if r != 1 else date_tmp
            (st6.cell(row=r, column=cnt_st6+1)).value = rat_lst[5] if r != 1 else date_tmp
        loguru.logger.success('Completion OK: Increase rate')
        sub.xls_wb_off(wb,path_xls)


        loguru.logger.info('>> STEP4. Calculate slope from  ... '+str(path_xls))
        wb = sub.xls_wb_on(path_xls)
        st1 = sub.xls_st_on(wb,0,'20ma'   , 7)
        st2 = sub.xls_st_on(wb,0,'Slope20',16)
        st3 = sub.xls_st_on(wb,0,'60ma'    ,9)
        st4 = sub.xls_st_on(wb,0,'Slope60',17)
        cnt_st2 = st2.max_column
        cnt_st4 = st4.max_column
        pgs3 = cls.ProgressBar(st1.max_row)
        for r in range(1,st1.max_row+1):
            if r != 1:
                val=sub.cal_slope_rate(st1,r)
                val2=sub.cal_slope_rate(st3,r)
            pgs3.update()
            (st2.cell(row=r, column=cnt_st2+1)).value = val if r != 1 else date_tmp
            (st4.cell(row=r, column=cnt_st4+1)).value = val2 if r != 1 else date_tmp
        loguru.logger.success('Completion OK: Slope value')
        sub.xls_wb_off(wb,path_xls)


        loguru.logger.info('>> STEP5. Calculate price vs avg line price ... '+str(path_xls))
        wb = sub.xls_wb_on(path_xls)
        st1 = sub.xls_st_on(wb,0,'Price', 0)
        st2 = sub.xls_st_on(wb,0,'20ma' , 7)
        st3 = sub.xls_st_on(wb,0,'20CMT',18)
        st4 = sub.xls_st_on(wb,0,'60ma' , 9)
        st5 = sub.xls_st_on(wb,0,'60CMT',19)

        cnt_st3 = st3.max_column
        cnt_st5 = st5.max_column
        pgs4 = cls.ProgressBar(st1.max_row)
        for r in range(1,st1.max_row+1):
            if r != 1 :
                cmt_tmp  = str(sub.cal_price_position(st1,st2,r,'20ma'))
                cmt_tmp2 = str(sub.cal_price_position(st1,st4,r,'60ma'))
            pgs4.update()
            (st3.cell(row=r, column=cnt_st3+1)).value = str(cmt_tmp)  if r != 1 else date_tmp
            (st5.cell(row=r, column=cnt_st5+1)).value = str(cmt_tmp2) if r != 1 else date_tmp
        loguru.logger.success('Completion OK: Price vs Avg line price relation')
        sub.xls_wb_off(wb,path_xls)


        loguru.logger.info('>> STEP6. Calculate the change in value in 3 days ... '+str(path_xls))
        wb = sub.xls_wb_on(path_xls)
        st1 = sub.xls_st_on(wb,0,'Value',2)
        st2 = sub.xls_st_on(wb,0,'Vrate',21)
        cnt_st1 = st1.max_column
        cnt_st2 = st2.max_column
        pgs5 = cls.ProgressBar(st1.max_row)
        for r in range(1,st1.max_row+1):
            if r != 1 : vrt_tmp = sub.cal_value_increase_rate(st1,r)
            pgs5.update()
            (st2.cell(row=r, column=cnt_st2+1)).value = vrt_tmp if r != 1 else date_tmp
        loguru.logger.success('Completion OK: Value rate')
        sub.xls_wb_off(wb,path_xls)


        loguru.logger.info('>> STEP7. Combine today data in the same sheet ... ')
        wb = sub.xls_wb_on(path_xls)
        wb_out = sub.xls_wb_on(path_xls_tdout)
        st_out = sub.xls_st_on(wb_out,0,'Today',0)
        st_out.delete_cols(1,30)
        tmp_clm = st_out.max_column
        st_dict = {
                    'Price':'股價', 'Volume':'成交量', 'Value':'成交值', 'Turnover':'周轉率',
                    '3ma':'3日線', '5ma':'5日線', '10ma':'10日線', '20ma':'20日線', '40ma':'40日線', '60ma':'60日線',
                    'Inc1':'1日漲幅', 'Inc3':'3日漲幅', 'Inc5':'5日漲幅', 'Inc10':'10日漲幅', 'Inc20':'20日漲幅', 'Inc60':'60日漲幅',
                    'Slope20':'月線斜率', 'Slope60':'季線斜率', '20CMT':'站上月線?', '60CMT':'站上季線?', 'Tangled':'均線糾結?', 'Vrate':'值增率', 'Force':'主力'
                }
        pgs6 = cls.ProgressBar(len(st_dict))
        for nam in wb.sheetnames:
            st = wb[nam]
            lst_tmp = []
            for i in range(st.max_row): lst_tmp.append(0)
            for r in range(1,st.max_row+1): lst_tmp[r-1] = ((st.cell(row=r , column=st.max_column)).value) if r != 1 else st_dict[nam]
            for r in range(1,st.max_row+1):
                (st_out.cell(row=r, column=tmp_clm+1)).font = Font(name='Calibri', size=12)
                (st_out.cell(row=r, column=tmp_clm+1)).value = lst_tmp[r-1]
            tmp_clm+=1
            pgs6.update()
        loguru.logger.success('Completion OK: Combination')
        sub.xls_wb_off(wb,path_xls)
        sub.xls_wb_off(wb_out,path_xls_tdout)


    end_time = time.time()
    loguru.logger.info('\n運算時間 : '+str(round((end_time-start_time),2))+'(S) >> '+str(round(round((end_time-start_time),2)/60,2))+'(min)')


    duration = 1000 # mS
    freq = 400      # Hz
    for i in range(5):
        if i%2 == 0 : winsound.Beep(freq, duration)
        time.sleep(0.2)

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
