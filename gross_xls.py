import loguru
import time
import datetime
import gross_sub as sub
import gross_display as dis
import gross_yahoo as ya


def cal_init():
    global TYP_VAL,TYP_TRN,date_tmp
    TYP_VAL=2.0
    TYP_TRN=1.0
    return 0



# ======        ======
# ======  main  ======
# ======        ======
if __name__ == '__main__':

    para1 = input('>> Capture data   1(Yes)/0(No) : ')
    para2 = input('>> Calculate data 1(Yes)/0(No) : ')
    loguru.logger.add( f'Stock_datalog_{datetime.date.today():%Y%m%d}.log', rotation='1 day', retention='7 days', level='DEBUG')
    TEST_M = 0
    DEMO_M = 0
    HIDAR_EXCEL = [int(para1),int(para2)]

    cal_init()
    sub.time_title()
    start_time = time.time()
    date_tmp = sub.get_stock_datetime()

    path_xls = 'C:\\Users\\JS Wang\\Desktop\\test\\tmp.xlsx'
    path_fin = 'C:\\Users\\JS Wang\\Desktop\\test\\gross_all_0115.txt'

    if( TEST_M == 1 ): ya.yahoo_stock_data()
    if( DEMO_M == 1 ): dis.display(path_xls)

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

        JS_TMP = [0,0]
        for line in lines:
            STK_NUM = int(line)
            JS_TMP = sub.parse_stock_data(sub.get_reqs_data(sub.get_stock_urls(str(STK_NUM))))
            STK_PRI.append(float(JS_TMP[0]))
            STK_VOL.append(float(JS_TMP[1]))
            STK_TRR.append(float(JS_TMP[2]))
            print(">> No."+str(STK_CNT) + "  ...  "+str(STK_NUM))
            sub.rand_on()
            STK_CNT+=1

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

        loguru.logger.success('Completion : Capture daily info.')
        wb.save(path_xls)

    if( HIDAR_EXCEL[1] == 1 ):

        loguru.logger.info('>> STEP2. Calculate average line price from  ... '+str(path_xls))
        wb = sub.xls_wb_on(path_xls)
        st0 = sub.xls_st_on(wb,0,'Price',0)
        st4 = sub.xls_st_on(wb,0,'3ma'  ,4)
        st5 = sub.xls_st_on(wb,0,'5ma'  ,5)
        st6 = sub.xls_st_on(wb,0,'10ma' ,6)
        st7 = sub.xls_st_on(wb,0,'20ma' ,7)

        cnt_c0 = st0.max_column
        cnt_c4 = st4.max_column
        cnt_c5 = st5.max_column
        cnt_c6 = st6.max_column
        cnt_c7 = st7.max_column

        for r in range(1,st0.max_row+1):
            day_lst=[ 3, 5,10,20]
            avg_lst=[ 0, 0, 0, 0]
            if r != 1 : avg_lst=sub.cal_avg_price(st0,day_lst,r)
            print(avg_lst)
            (st4.cell(row=r, column=cnt_c4+1)).value = avg_lst[0] if r != 1 else date_tmp
            (st5.cell(row=r, column=cnt_c5+1)).value = avg_lst[1] if r != 1 else date_tmp
            (st6.cell(row=r, column=cnt_c6+1)).value = avg_lst[2] if r != 1 else date_tmp
            (st7.cell(row=r, column=cnt_c7+1)).value = avg_lst[3] if r != 1 else date_tmp

        loguru.logger.success('Completion : Average line price')
        wb.save(path_xls)


        loguru.logger.info('>> STEP3. Calculate Increase rate from  ... '+str(path_xls))
        wb = sub.xls_wb_on(path_xls)
        st0 = sub.xls_st_on(wb,0,'Price',0)
        stA = sub.xls_st_on(wb,0, 'Inc1',10)
        stB = sub.xls_st_on(wb,0, 'Inc3',11)
        stC = sub.xls_st_on(wb,0, 'Inc5',12)
        stD = sub.xls_st_on(wb,0,'Inc10',13)
        stE = sub.xls_st_on(wb,0,'Inc20',14)

        cnt1d = stA.max_column
        cnt2d = stB.max_column
        cnt3d = stC.max_column
        cnt4d = stD.max_column
        cnt5d = stE.max_column

        for r in range(1,st0.max_row+1):
            day_lst=[ 1, 3, 5,10,20]
            rat_lst=[ 0, 0, 0, 0, 0]
            if r!= 1 : rat_lst=sub.cal_incr_rate(st0,day_lst,r)
            print(rat_lst)
            (stA.cell(row=r, column=cnt1d+1)).value = rat_lst[0] if r != 1 else date_tmp
            (stB.cell(row=r, column=cnt2d+1)).value = rat_lst[1] if r != 1 else date_tmp
            (stC.cell(row=r, column=cnt3d+1)).value = rat_lst[2] if r != 1 else date_tmp
            (stD.cell(row=r, column=cnt4d+1)).value = rat_lst[3] if r != 1 else date_tmp
            (stE.cell(row=r, column=cnt5d+1)).value = rat_lst[4] if r != 1 else date_tmp

        loguru.logger.success('Completion : Increase rate')
        wb.save(path_xls)

        loguru.logger.info('>> STEP4. Calculate slope from  ... '+str(path_xls))
        wb = sub.xls_wb_on(path_xls)
        stm1 = sub.xls_st_on(wb,0,'20ma'   , 7)
        stm2 = sub.xls_st_on(wb,0,'Slope20',15)
        cnt_m2 = stm2.max_column

        for r in range(1,stm1.max_row+1):
            if r != 1:
                val=sub.cal_slope_rate(stm1,r)
                print(str(val))
                (stm2.cell(row=r, column=cnt_m2+1)).value = val
            else:
                (stm2.cell(row=r, column=cnt_m2+1)).value = date_tmp
        loguru.logger.success('Completion : Slope value')
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
