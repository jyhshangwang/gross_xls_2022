import datetime
import fractions
import json
import os
import re

import loguru
import requests

class AfterHoursInfo:

    def __init__(self,code,name,totalShare,totalTurnover,\
        openPrice,highestPrice,lowestPrice,closePrice):

        self.Code = code
        self.Name = name
        self.TotalShare = self.checkNumber(totalShare)
        if self.TotalShare is not None: self.TotalShare = int(totalShare)

        self.TotalTurnover = self.checkNumber(totalTurnover)
        if self.TotalTurnover is not None: self.TotalTurnover = int(totalTurnover)

        self.OpenPrice = self.checkNumber(openPrice)
        if self.OpenPrice is not None: self.OpenPrice = fractions.Fraction(openPrice)

        self.HighestPrice = self.checkNumber(highestPrice)
        if self.HighestPrice is not None: self.HighestPrice = fractions.Fraction(highestPrice)

        self.LowestPrice = self.checkNumber(lowestPrice)
        if self.LowestPrice is not None: self.LowestPrice = fractions.Fraction(lowestPrice)

        self.ClosePrice = self.checkNumber(closePrice)
        if self.ClosePrice is not None: self.ClosePrice = fractions.Fraction(closePrice)

    def __repr__(self):
        totalShare = self.TotalShare
        if totalShare is not None:
            totalShare = f'{totalShare}'
        totalTurnover = self.TotalTurnover
        if totalTurnover is not None:
            totalTurnover = f'{totalTurnover}'
        openPrice = self.OpenPrice
        if openPrice is not None:
            openPrice = f'{float(openPrice):.2f}'
        highestPrice = self.HighestPrice
        if highestPrice is not None:
            highestPrice = f'{float(highestPrice):.2f}'
        lowestPrice = self.LowestPrice
        if lowestPrice is not None:
            lowestPrice = f'{float(lowestPrice):.2f}'
        closePrice = self.ClosePrice
        if closePrice is not None:
            closePrice = f'{float(closePrice):.2f}'
        '''
        return (
            f'class AfterHoursInfo {{ '
            f'Code={self.Code}, '
            f'Name={self.Name}, '
            f'TotalShare={totalShare}, '
            f'TotalTurnover={totalTurnover}, '
            f'OpenPrice={openPrice}, '
            f'HighestPrice={highestPrice}, '
            f'LowestPrice={lowestPrice}, '
            f'ClosePrice={closePrice} '
            f'}}'
        )
        '''
        return [self.Code,self.Name,self.TotalShare,self.TotalTurnover,\
            self.OpenPrice,self.HighestPrice,self.LowestPrice,self.ClosePrice]


    def checkNumber(self, value):
        if value == '--':
            return None
        else:
            return value


def main():

    fout = open(f'_stock_json_data_{datetime.date.today():%Y%m%d}.log','w')

    reqs = requests.get(
        f'https://www.twse.com.tw/exchangeReport/MI_INDEX?' +
        f'response=json&' +
        f'type=ALLBUT0999' +
        f'&date={datetime.date.today():%Y%m%d}'
    )

    if reqs.status_code != 200:
        loguru.logger.error('REQS: status code is not 200.')
        return
    else:
        loguru.logger.success('REQS: success.')
    
    afterHoursInfo = []

    body = reqs.json()
    print(body)
    print(type(body))
    stat = body['stat']
    if stat != 'OK':
        loguru.logger.error(f'REQS: body.stat error is {stat}.')
        return

    stocks = body['data9']
    #["5388","中磊","2,760,792","2,008","210,875,433","77.20","77.70","75.60","76.00","<p style= color:green>-<\u002fp>","0.10","75.90","4","76.00","18","18.31"]
    for stock in stocks:
        code = stock[0].strip()
        if re.match(r'^[1-9][0-9][0-9][0-9]$', code) is not None:
            name = stock[1].strip()
            totalShare = stock[2].replace(',', '').strip()
            totalTurnover = stock[4].replace(',', '').strip()
            openPrice = stock[5].replace(',', '').strip()
            highestPrice = stock[6].replace(',', '').strip()
            lowestPrice = stock[7].replace(',', '').strip()
            closePrice = stock[8].replace(',', '').strip()
            
            stk = AfterHoursInfo(
                code = code,
                name = name,
                totalShare = totalShare,
                totalTurnover = totalTurnover,
                openPrice = openPrice,
                highestPrice = highestPrice,
                lowestPrice = lowestPrice,
                closePrice = closePrice,
            )
            #afterHoursInfo.append(afterHoursInfo)
            afterHoursInfo.append(stk.__repr__())
    
    for i in range(len(afterHoursInfo)):
        print(str(afterHoursInfo[i]), file=fout)
    
    fout.close()



if __name__ == '__main__':

    loguru.logger.add(
        f'__afterHoursInfo__{datetime.date.today():%Y%m%d}.log',
        retention='7 days',
        rotation='1 day',
        level='DEBUG'
    )
    main()