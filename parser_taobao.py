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


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class taobao():
    def __init__(self):
        self.chrome_option = webdriver.ChromeOptions()
        self.chrome_option.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.chrome_option.add_argument("disable-blink-features=AutomationControlled")
        self.browser = webdriver.Chrome(os.path.join('.', "chromedriver.exe"), options=self.chrome_option)
        # 最大化視窗
        self.browser.maximize_window()
        self.browser.implicitly_wait(5)
        self.domain = 'http://www.taobao.com'
        self.action_chains = ActionChains(self.browser)

    def wait_element(self, tag, name):
        WebDriverWait(self.browser, 5).until(EC.presence_of_element_located((tag, name)))

    def login(self, username, password):
        while True:
            self.browser.get(self.domain)
            time.sleep(3)

            # 會xpath可以簡化這幾步
            self.browser.find_element(By.CLASS_NAME, 'h').click()
            self.wait_element(By.XPATH, '//iframe[@class="reg-iframe"]')

            self.browser.switch_to.frame(self.browser.find_element(By.XPATH, '//iframe[@class="reg-iframe"]'))
            self.browser.find_element(By.XPATH, '//*[@id="fm-login-id"]').send_keys(username)
            self.browser.find_element(By.XPATH, '//*[@id="fm-login-password"]').send_keys(password)

            time.sleep(1)

            # try:
            #     # 出現驗證碼，滑動驗證
            #     self.wait_element(By.XPATH, '//iframe[@id="baxia-dialog-content"]')
            #     self.browser.switch_to.frame(
            #         self.browser.find_element(By.XPATH, '//iframe[@id="baxia-dialog-content"]'))
            #     slider = self.browser.find_element(By.XPATH, "//span[contains(@class, 'btn_slide')]")
            #     if slider.is_displayed():
            #         # 拖拽滑塊
            #         self.action_chains.drag_and_drop_by_offset(slider, 258, 0).release().perform()
            #         time.sleep(0.5)
            #         # 釋放滑塊，相當於點選拖拽之後的釋放滑鼠
            #         # self.action_chains.release().perform()
            # except (NoSuchElementException, WebDriverException):
            #     logger.info('未出現登入驗證碼')

            # self.browser.find_element_by_class_name('password-login').click()
            self.browser.switch_to.parent_frame()
            time.sleep(1)
            self.browser.switch_to.frame(self.browser.find_element(By.XPATH, '//iframe[@class="reg-iframe"]'))
            time.sleep(1)
            self.browser.find_element(By.XPATH, '//*[@id="login-form"]/div[4]/button').click()
            nickname = self.get_nickname()
            if nickname:
                logger.info('登入成功，暱稱為:' + nickname)
                break
            logger.debug('登入錯誤，5s後繼續登入')
            time.sleep(5)

    def get_nickname(self):
        self.browser.get(self.domain)
        time.sleep(0.5)
        try:
            return self.browser.find_element(By.CLASS_NAME, 'site-nav-user').text
        except NoSuchElementException:
            return ''

    def switch_(self, query):
        url = f"https://s.taobao.com/search?q={query}&type=p&tmhkh5=&from=sea_1_searchbutton&catId=100&spm=a2141.241046-tw.searchbar.d_2_searchbox&bcoffset=1&ntoffset=1&p4ppushleft=2%2C48&s=0"
        self.browser.get(url)
        # 地區選全球
        time.sleep(2)
        button = self.browser.find_element(By.XPATH, '//*[@id="J_SiteNavBdL"]/li[1]/div[1]')
        button.click()
        self.wait_element(By.XPATH, '//*[@id="J_SiteNavRegionList"]/li[1]')
        self.browser.find_element(By.XPATH, '//*[@id="J_SiteNavRegionList"]/li[1]').click()
        time.sleep(3)
        # 銷量排序
        # self.wait_element(By.XPATH, '//*[@class="J_Ajax link  "]')
        # time.sleep(5)
        # self.browser.find_element(By.XPATH, '//*[@class="J_Ajax link  "]').click()

    def get_pages(self):
        try:
            get_page = self.browser.find_element(By.XPATH, '//*[@id="mainsrp-pager"]/div/div/div/div[1]')
            t_page = int(re.compile(r'(\d+)').search(get_page.text).group(1))
            return t_page
        except Exception as ex:
            return 1

    def parse_product(self, query, page_num):
        soup = BeautifulSoup(self.browser.page_source, 'lxml')
        comments = soup.find_all("div", class_="ctx-box J_MouseEneterLeave J_IconMoreNew")
        Infolist = []
        for i in comments:
            temp = [query]
            Item = i.find_all("div", class_="row row-2 title")  # 相關資訊
            href = i.find_all("a", class_="J_ClickStat", href=True)  # 連結資訊
            temp.append(Item[0].text.strip())  # 加入商品名稱
            shop = i.find_all("div", class_="row row-3 g-clearfix")
            for j in shop:
                a = j.find_all("span")
                temp.append(a[-1].text)  # 加入店鋪名稱
            address = i.find_all('div', class_='location')
            temp.append(address[0].text.strip())  # 加入店鋪所在地
            priceandnum = i.find_all("div", class_="row row-1 g-clearfix")
            for m in priceandnum:
                Y = m.find_all('div', class_='price g_price g_price-highlight')
                price = Y[0].text.strip()
                p_ = float(re.search(r'\d+[.,]\d+', price).group(0).replace(',', ''))
                temp.append(p_)  # 加入商品價格
                Num = m.find_all('div', class_='deal-cnt')
                temp.append(Num[0].text.strip())  # 加入購買人數
            temp.append(href[0]['href'])  # 加入連結
            Infolist.append(temp)
        # print(self.browser.page_source)
        _df = pd.DataFrame(Infolist, columns=['原料名稱', '商店標題', '店鋪名稱', '店鋪所在地', '商品價格', '購買人數', '連結'])
        if not _df.empty:
            print(f'原料名稱: {query} ------ 爬取頁數: {page_num}')
        else:
            print(f'原料名稱: {query} ------ 爬取失敗, 頁數: {page_num}')
        return _df

    def next_page(self):
        self.browser.find_element(By.XPATH, '//a[@trace="srp_bottom_pagedown"]').click()
        self.wait_element(By.XPATH, '//a[@trace="srp_bottom_pagedown"]')
        time.sleep(1)


if __name__ == '__main__':
    # 填入自己的使用者名稱，密碼
    username = 'xxx'
    password = 'xxx'
    tb = taobao()
    tb.login(username, password)
    search_name_list = ['赤藻糖醇', '赤藓糖醇', '乳糖醇', '山梨醇', '山梨糖醇', '異麥芽酮糖醇', '异麦芽酮糖醇', '木寡糖',
                        '低聚木糖', '果寡糖', '低聚果糖', '異麥芽寡糖', '低聚异麦芽糖', '木糖醇', '菊糖', '菊粉', '麥芽糖醇',
                        '麦芽糖醇', '阿拉伯膠', '阿拉伯胶', '難消化性麥芽糊精', '抗性糊精', '葡萄糖', '柑橘果膠', '果胶',
                        '刺槐豆膠', '槐豆胶', '麥芽糊精', '麦芽糊精', '麥芽糖', '麦芽糖', '關華豆膠', '瓜尔胶', '塔拉膠',
                        '刺云实胶']
    # test = ['山梨醇']

    for q in search_name_list:  # 要爬的關鍵字
        df_list = []
        tb.switch_(q)
        time.sleep(5)
        total_page = tb.get_pages()
        print('total_page', total_page)
        n = 1
        try:
            while n <= total_page:
                temp_parse_info = tb.parse_product(q, n)
                if not temp_parse_info.empty:
                    df_list.append(temp_parse_info)
                time.sleep(1)
                if n != total_page:
                    tb.next_page()
                n += 1
        except Exception as e:
            print(e)

        if df_list:
            df = pd.concat(df_list, ignore_index=True)
            today = datetime.datetime.now().date().strftime('%Y%m%d')
            df = df.drop_duplicates(subset=['連結'], ignore_index=True)
            df.to_excel(os.path.join('.', 'exported_data', f'{q}_淘寶原料產品_{today}.xlsx'), index=False)

        # print(tb.get_pages())
        # # combine data
        all_df = []
        for i in os.listdir(os.path.join('.', 'exported_data')):
            temp = pd.read_excel(os.path.join('.', 'exported_data', i))
            all_df.append(temp)

        all_df = pd.concat(all_df)
        all_ = all_df.drop_duplicates(subset=['連結'], ignore_index=True)
        path_ = os.path.join('.', 'exported_data', f"淘寶原料產品_all_{datetime.datetime.now().date().strftime('%Y%m%d')}.xlsx")
        _all_ = all_.drop_duplicates(subset=['商店標題', '店鋪名稱'], ignore_index=True)
        _all_.loc[:, '連結'] = _all_.loc[:, '連結'].apply(lambda x: 'https:' + x if 'https:' not in x else x)
        _all_.to_excel(path_, index=False)