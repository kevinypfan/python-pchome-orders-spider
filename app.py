from pchome_spider import PchomeSpider

import getpass

print('請輸入 pchome 帳號:')
username = input()

password = getpass.getpass('密碼:')

pchome = PchomeSpider(username, password)
pchome.first_login()
pchome.load_cookies()
pchome.init_requests_session()
pchome.get_all_orders()
pchome.conver_dataframe()
pchome.export_xlsx()