import os
import time

import collections
import requests
import random
import pyexcel
from bs4 import BeautifulSoup

# user_agent库：每次执行一次访问随机选取一个 user_agent，防止过于频繁访问被禁止
USER_AGENT_LIST = [
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1",
    "Mozilla/5.0 (X11; CrOS i686 2268.111.0) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1092.0 Safari/536.6",
    "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1090.0 Safari/536.6",
    "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; 360SE)",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.0 Safari/536.3",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24",
    "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24"]

headers = {'user-agent': random.choice(USER_AGENT_LIST),
           'Accept-Encoding': 'gzip, deflate',
           'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
           'Cache-Control': 'max-age=0',
           'Connection': 'keep-alive',
           'Host': 'b.zol.com.cn',
           'Accept-Language': 'en,zh-CN;q=0.9,zh;q=0.8,zh-TW;q=0.7,ja;q=0.6,la;q=0.5',
           'Referer': 'http://b.zol.com.cn/qy/all/list_51_31_0_1.html?curFloor=1&searchKey=&authType=&registerType=2',
           'Cookie': 'ip_ck=5sOD7vj0j7QuMTQyOTIzLjE2MDA3NjU3NzM%3D; lv=1603092765; vn=2; error_url=http%3A%2F%2Fb.zol.com.cn%2F; http_referer=http%3A%2F%2Fb.zol.com.cn%2F; errorURLHost=b.zol.com.cn; channelName=%D6%D0%B9%D8%B4%E5%D4%DA%CF%DF; categoryid=39; Adshow=0; questionnaire_pv=1603065601; Hm_lvt_ae5edc2bc4fc71370807f6187f0a2dd0=1600765777,1603093825; Hm_lpvt_ae5edc2bc4fc71370807f6187f0a2dd0=1603093825; z_pro_city=s_provice%3Dbeijing%26s_city%3Dbeijing; z_day=ixgo20%3D1'}

def start_craw(url):
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')

    agent_info_list = soup.find_all(class_='list-info')

    agent_list = []

    for a in agent_info_list:
        title = a.find('h3', class_='info-title').text.strip()
        agent_info = {'代理商名称': title}

        business_info = a.find_all('div', class_='list-info-business')

        for b in business_info:
            business_range = b.find('span', class_='business-range')
            business = b.find('span', class_='business')

            if business_range and business_range:
                sub = business_range.text.strip()
                value = business.text.strip().replace('\n', ',')

                if sub.find('经营范围') != -1:
                    agent_info.update({'经营范围': value})

                if sub.find('热门业务') != -1:
                    agent_info.update({'热门业务': value})

                if sub.find('联系人') != -1:
                    agent_info.update({'联系人': value})

                if sub.find('地址') != -1:
                    agent_info.update({'商家地址': value})

        agent_list.append(agent_info)

    return agent_list


def craw_book(register_type, book_name, from_page, to_page):
    if not os.path.isfile(book_name) or not os.access(book_name, os.R_OK):
        info = collections.OrderedDict({'代理商名称':'', '经营范围':'', '热门业务':'', '联系人':'', '商家地址':''})
        sheet = pyexcel.get_sheet(adict=info)
        sheet.save_as(book_name)

    sheet = pyexcel.get_sheet(file_name=book_name)

    for page in range(from_page, to_page + 1):
        print(f'processing page {page}...', end='')
        url = f'http://b.zol.com.cn/qy/all/list_51_0_0_{page}.html?curFloor=1&searchKey=&authType=&registerType={register_type}'
        info_list = start_craw(url)

        for info in info_list:
            sheet.row += [info['代理商名称'], info['经营范围'], info['热门业务'], info['联系人'], info['商家地址']]

        sheet.save_as(book_name)
        print('done')

        time.sleep(8)


if __name__ == "__main__":
    print('开始渠道商')
    craw_book('1', '渠道商.xlsx', 1, 147)
    print('渠道商结束')
    print('开始制造商')
    craw_book('2', '制造商.xlsx', 1, 1)
    print('制造商结束')
    print('开始服务商')
    craw_book('3', '服务商.xlsx', 1, 47)
    print('服务商结束')

    book1 = pyexcel.get_book(file_name='渠道商.xlsx')
    book2 = pyexcel.get_book(file_name='制造商.xlsx')
    book3 = pyexcel.get_book(file_name='服务商.xlsx')

    book = book1 + book2 + book3

    book.save_as('商.xls')