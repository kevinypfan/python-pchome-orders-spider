from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException
from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
import time
import sys
import json
import random
import requests
import math
import pandas as pd
from multiprocessing import Process



class PchomeSpider:
    def __init__(self, email, password):
        self.cookies = None
        self.email = email
        self.password = password
        self.requests = None
        self.order_ids = None
        self.orders = {}
        self.df = None

    def print_log(self, log):
        print(log, flush=True)
        sys.stdout.flush()

    def load_cookies(self):
        with open('vcyber.json') as f:
            self.cookies = json.load(f)

    def first_login(self):
        option = webdriver.ChromeOptions()
        # option.add_argument(f"user-agent={self.ua.google}")
        driver = webdriver.Chrome(chrome_options=option,
                                  executable_path='./chromedriver')
        self.base_url = 'https://ecvip.pchome.com.tw/login/v3/login.htm?rurl='
        driver.get("https://ecvip.pchome.com.tw/web/order/all")
        time.sleep(1)
        input_email = driver.find_element_by_id("loginAcc")
        input_password = driver.find_element_by_id("loginPwd")
        btn_login = driver.find_element_by_id("btnLogin")

        input_email.send_keys(self.email)
        input_password.send_keys(self.password)
        time.sleep(1)
        btn_login.click()
        time.sleep(8)

        self.cookies = driver.get_cookies()
        jsonCookies = json.dumps(self.cookies)
        with open('vcyber.json', 'w') as f:
            f.write(jsonCookies)

        driver.quit()
    
    def init_requests_session(self):
        user_agent = {'User-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36'}

        s = requests.Session()

        s.headers.update(user_agent)

        for cookie in self.cookies:
            s.cookies.set(cookie['name'], cookie['value'])
        
        self.requests = s
        
    def get_order_ids_by_page(self, current_page=1, row_per_page=20):
        order_ids_api = "https://ecvip.pchome.com.tw/ecapi/order/v2/index.php/core/order"
        id_set = set()
        params = {'site': 'ecshop', 'offset': (current_page - 1) * 20 + 1, 'limit': row_per_page}
        res = self.requests.get(order_ids_api, params=params)
        json = res.json()
        for row in json['Rows']:
            id_set.add(row['Id'])
        return list(id_set)
    
    def get_total_rows(self):
        order_ids_api = "https://ecvip.pchome.com.tw/ecapi/order/v2/index.php/core/order"
        params = {'site': 'ecshop'}
        res = self.requests.get(order_ids_api, params=params)
        json = res.json()
        return json['TotalRows']
    
    def get_all_order_ids(self):
        self.print_log('取得訂單編號中')
        
        row_per_page = 20
        
        id_list = []
        
        total_rows = self.get_total_rows()
        
        total_pages = math.ceil(total_rows / row_per_page)
        
        time.sleep(0.6)
        
        for i in range(0, total_pages):
            id_list = id_list + self.get_order_ids_by_page(i+1)
            time.sleep(0.6)
        
        self.order_ids = id_list
        return id_list
    
    def get_all_orders_info(self):
        self.print_log('取得訂單資訊中')
        row_per_page = 20

        ids_pgae_list = [self.order_ids[x:x+row_per_page] for x in range(0, len(self.order_ids),row_per_page)]
        for id_page in ids_pgae_list:
            order_lists_api = "https://ecvip.pchome.com.tw/fsapi/order/v1/main"
            params = {}
            ids_str = ','.join(id_page)
            params['id'] = ids_str
            res = self.requests.get(order_lists_api, params=params)
            json = res.json()
            for key in json:
                self.orders[key] = json[key]
            time.sleep(0.6)
                
    def get_all_order_prods(self):
        self.print_log('取得訂單商品列表中')
        for key in self.order_ids:
            order_details_api = "https://ecvip.pchome.com.tw/fsapi/order/v1/detail"
            params = {}
            params['id'] = key
            res = self.requests.get(order_details_api, params=params)
            json = res.json()
            products = json[key]['Detail']
            self.orders[key]['Products'] = products
            time.sleep(0.6)
            self.print_log(f'取得訂單商品列表中: {key}')
            
    def get_all_orders(self):
        self.get_all_order_ids()
        self.get_all_orders_info()
        self.get_all_order_prods()
    
    def conver_dataframe(self):
        columns = ['訂單編號', '日期', '時間', '訂單狀態', '總金額', '發票號碼', '付款方式']

        data=[] 
        for key in self.orders:
            data.append((key
                        ,self.orders[key]['OrderDate']
                        ,self.orders[key]['OrderTime']
                        ,self.orders[key]['OrderStatus']
                        ,self.orders[key]['Total']
                        ,self.orders[key]['InvoiceNo']
                        ,self.orders[key]['Payway'][0]['PayType']))
        df = pd.DataFrame(data=data,columns=columns)
        self.df = df
    
    def export_xlsx(self):
        self.df.to_excel (r'export_dataframe.xlsx', index = False, header=True)
        
