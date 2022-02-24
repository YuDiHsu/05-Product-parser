import datetime
import logging
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import os
from bs4 import BeautifulSoup
import re
import pandas as pd
from lxml import etree

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class IHerb():
    def __init__(self):
        self.chrome_option = webdriver.ChromeOptions()
        self.chrome_option.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.chrome_option.add_argument("disable-blink-features=AutomationControlled")
        self.browser = webdriver.Chrome(os.path.join('.', "chromedriver.exe"), options=self.chrome_option)
        # 最大化視窗
        self.browser.maximize_window()
        self.browser.implicitly_wait(5)
        self.domain = 'https://tw.iherb.com/'
        self.action_chains = ActionChains(self.browser)

    def wait_element(self, tag, name):
        WebDriverWait(self.browser, 5).until(EC.presence_of_element_located((tag, name)))

    def get_page(self):
        self.browser.get(self.domain)
        self.wait_element(By.XPATH, '//*[@id="gtm-no-shipping-content"]')
        self.browser.find_element(By.XPATH, '//*[@id="gtm-no-shipping-popup-close"]/i').click()
        time.sleep(1)
        # self.browser.find_element(By.CLASS_NAME, 'icon welcome-mat-module-close').click()
        # time.sleep(1)

    def get_max_page_num(self):
        soup = BeautifulSoup(self.browser.page_source, 'html.parser')
        # test = soup.find_all("span", class_="circle")
        time.sleep(2)
        try:
            test = soup.find_all("a", class_='pagination-link')
            page_list = []
            for i in test:
                page_num = i.text.strip().split('\n\t\t\n\t\t')[0]
                if page_num:
                    page_list.append(int(page_num))
            print(page_list)
            return max(page_list)
        except Exception as ex:
            print(ex)
            return 1

    def search(self, search_word):
        self.wait_element(By.XPATH, '//*[@id="txtSearch"]')
        self.browser.find_element(By.XPATH, '//*[@id="txtSearch"]').send_keys(search_word)
        self.browser.find_element(By.XPATH, '//*[@id="searchBtn"]').click()
        try:
            self.wait_element(By.CLASS_NAME, 'sub-header')
        except Exception as e:
            print(e)
        time.sleep(2)

    def parse_product(self, search_word, page_num):
        data_list = []
        for p in range(page_num):
            url_ = f'https://tw.iherb.com/search?kw={search_word}&p={p+1}'
            self.browser.get(url_)
            self.wait_element(By.CLASS_NAME, 'sub-header')
            time.sleep(1)
            soup = BeautifulSoup(self.browser.page_source, 'html.parser')
            comments = soup.find_all("div", class_="absolute-link-wrapper")
            for i in comments:
                # print(i.a.attrs)
                # print('-'*100)
                if 'title' in i.a.attrs and 'data-ga-discount-price' in i.a.attrs:
                    title_ = i.a['title']
                    company = title_.split(', ')[0].strip()
                    product = title_.split(', ')[1].strip()
                    price = i.a['data-ga-discount-price']
                    link = i.a['href']
                    data_list.append([search_word, company, product, price, link])
                    print(f"爬取第{p+1}頁中--總頁數:{page_num}--{[search_word, company, product, price, link]}")
            print(f"第{p+1}頁爬取結束--總頁數:{page_num}")

        df = pd.DataFrame(data_list, columns=['原料名稱', '公司名', '商品名', '商品價格', '連結'], index=None)
        return df


if __name__ == '__main__':
    ih = IHerb()
    ih.get_page()
    all_data_list = []
    search_name_list = ['赤藻糖醇', '赤藓糖醇', '乳糖醇', '山梨醇', '山梨糖醇', '異麥芽酮糖醇', '异麦芽酮糖醇', '木寡糖',
                        '低聚木糖', '果寡糖', '低聚果糖', '異麥芽寡糖', '低聚异麦芽糖', '木糖醇', '菊糖', '菊粉', '麥芽糖醇',
                        '麦芽糖醇', '阿拉伯膠', '阿拉伯胶', '難消化性麥芽糊精', '抗性糊精', '葡萄糖', '柑橘果膠', '果胶',
                        '刺槐豆膠', '槐豆胶', '麥芽糊精', '麦芽糊精', '麥芽糖', '麦芽糖', '關華豆膠', '瓜尔胶', '塔拉膠',
                        '刺云实胶']
    # rest_search = ['難消化性麥芽糊精', '刺槐豆膠', '槐豆胶', '麥芽糊精', '麦芽糊精', '麥芽糖', '麦芽糖', '關華豆膠',
    #                '瓜尔胶', '塔拉膠', '刺云实胶']
    # ih.parse_product('赤藻糖醇')
    # test = ['山梨糖醇']
    today = datetime.datetime.now().date().strftime('%Y%m%d')
    for sr in search_name_list:
        ih.search(sr)
        max_page = ih.get_max_page_num()
        time.sleep(2)
        df = ih.parse_product(sr, max_page)
        df.to_excel(os.path.join('.', 'exported_data_iherb', 'temp', f'{sr}_iherb原料產品_{today}.xlsx'), index=False)

    all_df = []
    for i in os.listdir(os.path.join('.', 'exported_data_iherb', 'temp')):
        temp = pd.read_excel(os.path.join('.', 'exported_data_iherb', 'temp', i))
        all_df.append(temp)

    all_df = pd.concat(all_df)
    all_ = all_df.drop_duplicates(subset=['連結'], ignore_index=True)
    path_ = os.path.join('.', 'exported_data_iherb',
                         f"iherb原料產品_all_{datetime.datetime.now().date().strftime('%Y%m%d')}")
    _all_ = all_.drop_duplicates(subset=['公司名', '商品名'], ignore_index=True)

    _all_.to_excel(path_, index=False)
