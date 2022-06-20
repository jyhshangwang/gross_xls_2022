

class ProgressBar:
    bar_str_fmt = '\r>> Progress: |{}{}| {:.2%}  ({}/{})'
    cnt = 0

    def __init__(self, total):
        self.Stk_cnt = total
        self.Bar_len = 40

    def update(self, step=1):
        all = self.Stk_cnt
        self.cnt += step
        done_cnt = int((self.cnt/all)*self.Bar_len)
        undone_cnt = self.Bar_len - done_cnt

        progress = self.bar_str_fmt.format( '█' * done_cnt, ' ' * undone_cnt, self.cnt/all, self.cnt, all)
        print(progress, end='')

        percent = self.cnt/all
        if   percent == 1: print('\n')
        elif percent >= 1: print('')

        if self.cnt == self.Stk_cnt: self.cnt = 0


class ProportionDailyInfo:

    def __init__(self,op_price,hi_price,lo_price,td_price,up_down,hi_price_1y,lo_price_1y,pe_ratio,mx_volume_1y,mi_volume_1y,td_volume,incr_year,stk_count,Rev_rat_cmt):
        self.OP_Price = op_price # 開盤價
        self.HI_Price = hi_price # 盤中最高價
        self.LO_Price = lo_price # 盤中最低價
        self.TD_Price = td_price # 收盤價
        self.Up_Down = up_down # 今日漲跌
        self.YD_Price = round((td_price-up_down),2) # 昨日收盤價
        self.TD_Volume = td_volume # 今日成交量
        self.TurnOver = round((td_volume/stk_count/100),2) # 今日周轉率
        self.TD_Value = round((td_price*td_volume/100000),2) # 今日成交值
        self.PE_Ratio = pe_ratio # 今日本益比
        self.HI_Price_1y = hi_price_1y # 1年內最高價
        self.LO_Price_1y = lo_price_1y # 1年內最低價
        self.MX_Volume_1y = mx_volume_1y # 1年內最大量
        self.MI_Volume_1y = mi_volume_1y # 1年內最小量
        self.INC_Year = incr_year # 今年以來漲幅率
        self.STK_Count = stk_count # 股本(億)
        self.REV_RAT_CMT = Rev_rat_cmt # 營收比重說明
    def __repr__(self):
        return (
            f'{self.REV_RAT_CMT};'
            f'{self.PE_Ratio};'
            f'{self.MX_Volume_1y};'
            f'{self.MI_Volume_1y};'
            f'{self.TD_Volume};'
            f'{self.TurnOver};'
            f'{self.STK_Count};'
            f'{self.Up_Down};'
            f'{self.HI_Price_1y};'
            f'{self.LO_Price_1y};'
            f'{self.TD_Price};'
            f'{self.YD_Price};'
        )

class ProportionRevenueInfo:

    def __init__(self,month,revenue,mom,yoy,total_yoy):
        self.Month = month
        self.Revenue = revenue
        self.Mom = mom
        self.Yoy = yoy
        self.Tyoy = total_yoy
    def __repr__(self):
        return (
        f'{self.Month};'
        f'{self.Revenue};'
        f'{self.Mom};'
        f'{self.Yoy};'
        f'{self.Tyoy};'
        )
    def get_revenue(self):
        return float((self.Revenue).replace(',',''))
    def get_revenue_100m(self):
        return float((self.Revenue).replace(',',''))/100000
    def get_yoyrate(self):
        return self.Yoy

