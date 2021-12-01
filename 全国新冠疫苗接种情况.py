from openpyxl import Workbook
import re, time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

"""
* 1.python先安装 openpyxl selenium
* 2.安装Firefox浏览器 https://www.firefox.com.cn/download/
"""
s = Service(r"geckodriver.exe")
url = r'http://www.nhc.gov.cn'
browser = webdriver.Firefox(service=s)
pages = 2


def save_as_xlsx(header, data: '表格数据', path: '文件路径'):
    wb = Workbook()
    ws = wb.active
    ws.append(header)
    for col in data:
        ws.append(list(col))
    wb.save(path)
    wb.close()


def get_catalogue(page=1):
    __url__ = f'{url}/xcs/yqjzqk/list_gzbd.shtml' if page == 1 else f'{url}/xcs/yqjzqk/list_gzbd_{page}.shtml'
    browser.get(__url__)
    regex = r"href=\"(/.*\.shtml)\" target=\"_blank\""
    time.sleep(3)
    url_li = re.findall(regex, browser.page_source)
    print(url_li)
    return url_li


def get_page(url_li):
    regex = r"(\d+年\d+月\d+日).*疫苗(\d+.\d+)万剂次"
    _li_ = []
    for e in url_li:
        time.sleep(1)
        __url__ = f'{url}{e}'
        browser.get(__url__)
        time.sleep(3)
        res = re.findall(regex, browser.page_source, re.MULTILINE)
        if len(res) > 0:
            _li_.append(res[0])
            print(res[0])
    return _li_


if __name__ == '__main__':
    all_li = []
    for i in range(1, pages + 1):
        urls = get_catalogue(i)
        li = get_page(urls)
        all_li.extend(li)
        print(li)
    print(all_li)
    save_as_xlsx(header=['日期', '万剂次'], data=all_li, path='全国新冠病毒疫苗接种情况1.0.xlsx')
    browser.quit()
