from asyncore import read
import time
import loguru
import pyquery
import requests
import chardet
import csv
import datetime
import os
import random
import requests
import requests.exceptions



def getProxiesFromFreeProxyList():

    proxies = []

    url = 'https://free-proxy-list.net/'
    loguru.logger.debug(f'getProxiesFromFreeProxyList: {url}')
    loguru.logger.warning(f'getProxiesFromFreeProxyList: downloading...')
    reqs = requests.get(url)
    if reqs.status_code != 200:
        loguru.logger.error(f'getProxiesFromFreeProxyList: status code is not 200.')
        return
    loguru.logger.success(f'getProxiesFromFreeProxyList: downloaded.')
    det = chardet.detect(reqs.content)
    txt = None
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
    if txt is None: return

    d = pyquery.PyQuery(txt)
    trs = list(d('table#proxylisttable > tbody > tr').items())
    loguru.logger.warning(f'getProxiesFromFreeProxyList: scanning...')

    for tr in trs:
        tds = list(tr('td').items())
        ip = tds[0].text().strip()
        port = tds[1].text().strip()
        proxy = f'{ip}:{port}'
        proxies.append(proxy)
    loguru.logger.success(f'getProxiesFromFreeProxyList: scanned.')
    loguru.logger.debug(f'getProxiesFromFreeProxyList: {len(proxies)} proxies is found.')

    return proxies



now = datetime.datetime.now()
proxies = []

def getProxy():
    global proxies
    if len(proxies) == 0:
        getProxies()
    proxy = random.choice(proxies)
    loguru.logger.debug(f'getProxy: {proxy}')
    proxies.remove(proxy)
    loguru.logger.debug(f'getProxy: {len(proxies)} proxies is unused.')
    return proxy

def reqProxies(hour):
    global proxies

    proxies = proxies + getProxiesFromFreeProxyList()
    proxies = list(dict.fromkeys(proxies))
    loguru.logger.debug(f'reqProxies: {len(proxies)} proxies is found.')

def getProxies():
    global proxies
    hour = f'{now:%Y%m%d%H}'
    filename = f'proxies-{hour}.csv'
    filepath = f'{filename}'
    if os.path.isfile(filepath):
        loguru.logger.info(f'getProxies: {filename} exists.')
        loguru.logger.warning(f'getProxies: {filename} is loading...')
        with open(filepath, 'r', newline='', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                proxy = row['Proxy']
                proxies.append(proxy)
        loguru.logger.success(f'getProxies: {filename} is loaded.')
    else:
        loguru.logger.info(f'getProxies: {filename} does not exist.')
        reqProxies(hour)
        loguru.logger.warning(f'getProxies: {filename} is saving...')
        with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow([
                'Proxy'
            ])
            for proxy in proxies:
                writer.writerow([
                    proxy
                ])
        loguru.logger.success(f'getProxies: {filename} is saved.')




proxy = None

def testRequest():
    global proxy
    while True:

        if proxy is None:
            proxy = getProxy()
        try:
            url = f'https://www.google.com/'
            loguru.logger.info(f'testRequest: url is {url}')
            loguru.logger.warning(f'testRequest: downloading...')
            response = requests.get(
                url,
                # 指定 HTTPS 代理資訊
                proxies={
                    'https': f'https://{proxy}'
                },
                # 指定連限逾時限制
                timeout=5
            )
            if response.status_code != 200:
                loguru.logger.debug(f'testRequest: status code is not 200.')
                # 請求發生錯誤，清除代理資訊，繼續下個迴圈
                proxy = None
                continue
            loguru.logger.success(f'testRequest: downloaded.')
        # 發生以下各種例外時，清除代理資訊，繼續下個迴圈
        except requests.exceptions.ConnectionError:
            loguru.logger.error(f'testRequest: proxy({proxy}) is not working (connection error).')
            proxy = None
            continue
        except requests.exceptions.ConnectTimeout:
            loguru.logger.error(f'testRequest: proxy({proxy}) is not working (connect timeout).')
            proxy = None
            continue
        except requests.exceptions.ProxyError:
            loguru.logger.error(f'testRequest: proxy({proxy}) is not working (proxy error).')
            proxy = None
            continue
        except requests.exceptions.SSLError:
            loguru.logger.error(f'testRequest: proxy({proxy}) is not working (ssl error).')
            proxy = None
            continue
        except Exception as e:
            loguru.logger.error(f'testRequest: proxy({proxy}) is not working.')
            loguru.logger.error(e)
            proxy = None
            continue

        break


















