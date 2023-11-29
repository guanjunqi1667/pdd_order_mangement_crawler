# -*- coding: utf-8 -*-
import os
# fix issue: urlopen error [SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed.
import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import time
from datetime import datetime, timezone, timedelta
import warnings
warnings.filterwarnings("ignore")
import xlwt
import msvcrt
import random

url = 'https://mms.pinduoduo.com/orders/list'
driver_path='C:/chromedriver_win32/chromedriver.exe'

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

################################### 设置selenium #############################################
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-automation'])
options.add_argument("--disable-blink-features=AutomationControlled")
service = Service(driver_path)

driver = webdriver.Chrome(options=options) # 使用Chrome浏览器
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
                    Object.defineProperty(navigator, 'webdriver', {
                      get: () => undefined
                    })
                  """
    })


###################################   设置表格   #############################################
global _Workbook
global order_time
_Workbook = object()
_Workbook = xlwt.Workbook()

sheet = _Workbook.add_sheet("待发货信息")
header_font = xlwt.Font()
header_font.name = "Arial"
header_font.bold = True
header_style = xlwt.XFStyle()
header_style.font = header_font
sheet.write(0, 0, "订单编号", header_style)
sheet.write(0, 1, "收件人", header_style)
sheet.write(0, 2, "手机", header_style)
sheet.write(0, 3, "地址", header_style)
sheet.write(0, 4, "商品系列名称", header_style)
sheet.write(0, 5, "商品名称", header_style)
sheet.write(0, 6, "发货数量", header_style)
sheet.write(0, 7, "备注", header_style)
sheet.write(0, 8, "逾期时间", header_style)

row_num = 1
timesleep = 0.6
days_selected = 15
page_number = 20
###############################################################################################


#print("当前网页: {}".format(driver.current_url))

######################################### 确定订单数量 #########################################
def get_ordernumber():
        underline = WebDriverWait(driver, 600).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, "PGT_totalText_5-92-0"))
        )
        underline = underline.pop()
        underline_text = underline.text
        underline_text_list = underline_text.split(" ")
        order_number = int(underline_text_list[1])
        return order_number

###################################### 自动翻开电话 #######################################
def phonenumber_check(_order_info_list_,_timesleep_):
    for order_info, title in zip(_order_info_list_[1::2],_order_info_list_[::2]):
        if '审核中' not in title.text:
            if '快递停运' in title.text:
                order_remaining_time = title.text.split('\n')[-2][37:-1].rstrip(' 后将逾期发')
            else:
                order_remaining_time = title.text.split('\n')[-2][21:-1].rstrip(' 后将逾期发')
            try:
                order_remaining_time = datetime.strptime(order_remaining_time,'%d天%H时%M分%S秒')
            except:
                    order_remaining_time = datetime.strptime(order_remaining_time,'%H时%M分%S秒')

            #order_time = datetime.strptime(title.text.split('\n')[-2][5:21], '%Y-%m-%d %H:%M')
            if order_remaining_time.day <= days_selected:
                try:
                    check_user_info_btn = WebDriverWait(order_info, 600).until(
                        EC.presence_of_element_located((By.LINK_TEXT, "查看"))
                    )
                    driver.execute_script("arguments[0].scrollIntoView();", check_user_info_btn)
                    driver.execute_script("arguments[0].click();", check_user_info_btn)
                    time.sleep(_timesleep_/2*random.randint(1,3))
                    check_phone_number_btn = WebDriverWait(order_info, 600).until(
                        EC.presence_of_element_located((By.LINK_TEXT, "查看手机号"))
                    )
                    driver.execute_script("arguments[0].click();", check_phone_number_btn)
                    time.sleep(_timesleep_/2*random.randint(1,3))

                except:
                    print('订单信息获取错误，请手动点击查看，完成操作后任意键继续')
                    msvcrt.getch()
                    print('程序继续执行... ...')
                    break
            else:
                    print('订单时间超出选择范围')
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(_timesleep_/2*random.randint(1,3))
                    return 0  

        else:
            continue
        time.sleep(_timesleep_/2*random.randint(1,3))
    return 1

        

    

###################################### 自动获取信息 #######################################
def get_infos(_order_info_list_,_row_num_):
    for order_info, title in zip(_order_info_list_[1::2],_order_info_list_[::2]):
        try:
            if '审核中'not in title.text:
                if '快递停运' in title.text:
                    order_remaining_time = title.text.split('\n')[-2][37:-1].rstrip(' 后将逾期发')
                else:
                    order_remaining_time = title.text.split('\n')[-2][21:-1].rstrip(' 后将逾期发')
                order_remaining_time_str = order_remaining_time
                try:
                    order_remaining_time = datetime.strptime(order_remaining_time,'%d天%H时%M分%S秒')
                except:
                    order_remaining_time = datetime.strptime(order_remaining_time,'%H时%M分%S秒')

                #order_time = datetime.strptime(title.text.split('\n')[-2][5:21], '%Y-%m-%d %H:%M')

                if order_remaining_time.day <= days_selected:
                    info_block = order_info.find_elements(By.TAG_NAME, "td")

                    product_serie_name = info_block[0].text.split('\n')[0]
                    product_name = info_block[0].text.split('\n').pop()
                    qty = info_block[2].text
                    usr_name = info_block[5].text.split('\n')[0]

                    if '隐私号' in info_block[5].text:
                        usr_phonenumber = info_block[5].text.split('\n')[3]
                    else:
                        usr_phonenumber = info_block[5].text.split('\n')[-2]
                    usr_addresse = info_block[5].text.split('\n')[-1]

                    id = title.text.split('\n')[0].lstrip("订单编号：")
                    if '有备注' in title.text:
                        note = info_block[7].text.split('\n')[-2].lstrip('用户备注:')
                        sheet.write(_row_num_, 7, note.strip())

                    sheet.write(_row_num_, 0, id.strip())
                    sheet.write(_row_num_, 1, usr_name.strip())
                    sheet.write(_row_num_, 2, usr_phonenumber.strip())
                    sheet.write(_row_num_, 3, usr_addresse.strip())
                    sheet.write(_row_num_, 4, product_serie_name.strip())
                    sheet.write(_row_num_, 5, product_name.strip())
                    sheet.write(_row_num_, 6, qty.strip())
                    sheet.write(_row_num_, 8, order_remaining_time_str.strip() + '后逾期')

                    _row_num_ += 1
                elif order_remaining_time.day > days_selected:
                    break
            else:
                continue
        except:
            print('程序意外退出，请重试')
            break

    return _row_num_

###################################### get_next #######################################
def get_next(_order_number,_page_number):
    if _order_number > _page_number:
        try:
            next = WebDriverWait(driver, 600).until(
                    EC.presence_of_all_elements_located((By.CLASS_NAME, "PGT_next_5-92-0"))
                )
            next[0].click()
            time.sleep(timesleep)
        except:
            print('Error: Can not locate Button_next_page, please try again')
###################################### main program #######################################

###################################### 启动selenium #######################################
driver.get(url)
driver.maximize_window()
expected_url = "https://mms.pinduoduo.com/orders/list"

################################## 等待登陆成功 ###############################################
WebDriverWait(driver, 6000).until(EC.url_contains(expected_url))

time_zone = timezone(timedelta(hours=8))              
dt_now = datetime.now(time_zone).replace(tzinfo=None)              # 获取当前时间的时间戳

#try:
#    warn1 = WebDriverWait(driver, 600).until(
#                EC.presence_of_all_elements_located((By.LINK_TEXT, "暂不处理"))
#            )
#    warn1.click()
#except:
#    warn1 = 'No Warning'

order_number = get_ordernumber()

while True:    
    #print('点击任意键继续')
    #msvcrt.getch()
    #print('继续执行... ...')

    order_table = WebDriverWait(driver, 600).until(
            EC.presence_of_element_located((By.ID, "order-content"))
        )

    order_info_list = WebDriverWait(driver, 600).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, "tr"))
        )
    order_info_list.pop(0)
    time.sleep(timesleep*random.randint(1,3))

    check_continue = phonenumber_check(order_info_list,timesleep)

    print('点击任意键继续')
    msvcrt.getch()
    print('继续执行... ...')

    row_num = get_infos(order_info_list,row_num)



    if check_continue and order_number > page_number:
        get_next(order_number,page_number)
        order_number -= page_number
    else:
        break


driver.close()
driver.quit()

if type(_Workbook) is not object:
    _Workbook.save("{}/Desktop/待发货表-{}.xls".format(
        os.path.expanduser("~"),
        time.strftime("%Y-%m-%d %H-%M-%S",time.localtime())
    ))
