from cmath import log
import datetime
import requests
import loguru
import pyquery
import chardet
import os

class Propotion:
    def __init__(self,month,revenue,mom,yoy,tyoy):
        self.Month = month
        self.Revenue = revenue
        self.Mom = mom
        self.Yoy = yoy
        self.Tyoy = tyoy
    def __repr__(self): #return (f'{{'f'Month={self.Month},'f'Revenue={self.Revenue},'f'mom={self.Mom},'f'yoy={self.Yoy},'f'tyoy={self.Tyoy},'f'}}')
        return f'{self.Month};{self.Revenue};{self.Mom};{self.Yoy};{self.Tyoy}'

@loguru.logger.catch
def main():
    reqs = requests.get('https://dj.mybank.com.tw/z/zc/zch/zch_3006.djhtm')
    if reqs.status_code != 200:
        loguru.logger.error('REQS: status code is not 200')
    loguru.logger.success('REQS: success')

    txt = None
    det = chardet.detect(reqs.content)
    try:
        if det['confidence'] > 0.5:
            if det['encoding'] == 'big-5':
                txt = reqs.content.decode('big5')
            else:
                txt = reqs.content.decode(det['encoding'])
        else:
            txt = reqs.content.decode('utf-8')
    except Exception as e:
        loguru.logger.error(e)
    #try代碼塊出錯則會創建Exception類(class)對象，對象名為e，e中封裝了出錯的錯誤訊息

    if txt is None: return
    
    #loguru.logger.info(txt)

    proportions = []

    d = pyquery.PyQuery(txt)
    trs = list(d('table tr').items())
    trs = trs[8:33]
    for tr in trs:
        tds = list(tr('td').items())
        code = tds[1].text().strip()
        if code != '':
            month = tds[0].text().strip()
            reven = tds[1].text().strip()
            mom   = tds[2].text().strip()
            yoy   = tds[4].text().strip()
            tyoy  = tds[6].text().strip()
            proportions.append(Propotion(month,reven,mom,yoy,tyoy))

    #proportions.sort(key=lambda proportion: proportion.Month)
    #loguru.logger.info(proportions)

    message = os.linesep.join([str(proportion) for proportion in proportions])
    loguru.logger.info('REVENUE' + os.linesep + message)


if __name__ == '__main__':
    loguru.logger.add(
        f'stock_datalog_{datetime.datetime.today():%Y%m%d}.log',
        rotation='1 day',
        retention='7 days',
        level='DEBUG'
    )
    main()