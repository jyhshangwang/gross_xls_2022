import time

import loguru
import pyquery
import requests

def getProxiesFromFreeProxyList():
    proxies = []
    url = 'https://free-proxy-list.net/'
    loguru.logger.debug(f'getProxiesFromFreeProxyList: {url}')
    loguru.logger.warning(f'getProxiesFromFreeProxyList: downloading...')
    response = requests.get(url)
    if response.status_code != 200:
        loguru.logger.debug(f'getProxiesFromFreeProxyList: status code is not 200')
        return
    loguru.logger.success(f'getProxiesFromFreeProxyList: downloaded.')
    d = pyquery.PyQuery(response.text)
    trs = list(d('table#proxylisttable > tbody > tr').items())
    loguru.logger.warning(f'getProxiesFromFreeProxyList: scanning...')
    for tr in trs:
        # 取出所有資料格
        tds = list(tr('td').items())
        # 取出 IP 欄位值
        ip = tds[0].text().strip()
        # 取出 Port 欄位值
        port = tds[1].text().strip()
        # 組合 IP 代理
        proxy = f'{ip}:{port}'
        proxies.append(proxy)
    loguru.logger.success(f'getProxiesFromFreeProxyList: scanned.')
    loguru.logger.debug(f'getProxiesFromFreeProxyList: {len(proxies)} proxies is found.')
    return proxies