from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from collections import OrderedDict
import time
import xlrd
import sys
import getopt
opts,args = getopt.getopt(sys.argv[1:],"hu:d:o:")

from selenium.webdriver.support.wait import WebDriverWait
browser = webdriver.Chrome()



def login():
    # 登陆应该提出去，不然每次都重新登录一次
    # 打开药监局网址
    browser.get("http://58.16.67.230:8001/index.html?user=ps")
    # 手动输入账户密码：[52P0168] [973844]
    login = """
    ********************************************************
    ****												****
    ****  请先在弹出来的网页中登录，再进行以下操作！！！    ****
    ****  账户 52P0168									****
    ****  密码 973844									****
    ****											    ****
    ********************************************************
    """
    print(login)


collapse = []
medicines = OrderedDict()

packId=0
uuid=[]
def open():
    global collapse,uuid
    global medicines
    global packId
    uuid = []
    collapse = []
    medicines = OrderedDict()
    # 输入订单包号
    if packId == 0 or packId == None:
        packId = input("输入订单包号：")
    url = "http://58.16.67.230:8001/orderDispatch/toOrderDispatchDetail.html?orderPackId=" + packId
    browser.get(url)
    time.sleep(0.5)
    # “采购数”展开的xpath = //*[@id="'+packId+'00001"]/td[2]/div
    # “药品名”展开的xpath = //*[@id="'+packId+'00001"]/td[9]
    # find_elements_by_css_selector("td[aria-describedby='gridlist_orderCode']")
    ids = browser.find_elements_by_css_selector("td[aria-describedby='gridlist_orderCode']")
    #print(ids)
    for i in ids:
        uid = i.get_attribute("title")
        uuid.append(i.get_attribute("title"))
        collapse.append('//*[@id="' + uid + '"]/td[2]/div')
        medicines['//*[@id="' + uid + '"]/td[9]'] = '//*[@id="' + uid + '"]/td[6]'



name_date = {}
name_number = {}
workbook = xlrd.open_workbook(r'.\\excel.xls')
sheet = workbook.sheet_by_index(0)
rows = sheet.nrows

for row in range(1,rows):
    row = sheet.row_values(row)
    name_date[row[7]] = row[-11]
    name_number[row[7]] = row[-13]
    # print(row[4],row[11],row[12])

    
def match(origin_name,name_number,name_date):
    #print(origin_name)
    for n in name_number.keys():
        #if re.match(n,origin_name):
        if origin_name in n:
            print('匹配到：',origin_name, '-----', name_number[n], '-------', name_date[n])
            print('*' * 40)
            return [name_number[n],name_date[n]]
    return []


# 以上数据都准备完毕 开始填入
# 批次 ：  //*[@id="0"]/td[3]/input     配送数量： //*[@id="0"]/td[4]/input  有效期：  //*[@id="0"]/td[5]/input

import random
def fill():
    global packId,uuid
    global name_number,name_date
    for index in range(0,len(collapse)):
        sno = "#subTable_"+uuid[index]+" > tbody > tr[id='0'] > td:nth-child(3) > input"
        number = "#subTable_"+uuid[index]+" > tbody > tr[id='0'] > td:nth-child(4) > input"
        date = "#subTable_"+uuid[index]+" > tbody > tr[id='0'] > td:nth-child(5) > input"
        k = [k for k in medicines.keys()][index]
        x = collapse[index]
        #time.sleep(1.8)
        #print(uuid[index])
        if index == 0 :
            browser.find_element_by_xpath(x).click()
        browser.find_element_by_xpath(x).click()
        
        #time.sleep(1.8)
        #WebDriverWait(browser, 3).until(lambda driver: driver.find_element_by_xpath(x))
        js = "$(\"" + date + "\").attr('readonly',false)"
        browser.execute_script(js)
        js = "$(\"" + date + "\").attr('onfocus',false)"
        browser.execute_script(js)
        v = medicines[k]
        origin_name = browser.find_element_by_xpath(k).text.strip()
        #print('origin:', origin_name)
        buy_number = browser.find_element_by_xpath(v).text
        m = match(origin_name,name_number,name_date)
        if len(m)>0:
            browser.find_element_by_css_selector(sno).send_keys(m[0])
            browser.find_element_by_css_selector(number).send_keys(buy_number)
            browser.find_element_by_css_selector(date).send_keys(m[1])
        else:
            if len(origin_name) < 4:
                name = origin_name[0:3]
            else:
                name = origin_name[0:4]
            m = match(name,name_number,name_date)
            if len(m) > 0:
                browser.find_element_by_css_selector(sno).send_keys(m[0])
                browser.find_element_by_css_selector(number).send_keys(buy_number)
                browser.find_element_by_css_selector(date).send_keys(m[1])
                # break
            else:
                idx = len(origin_name) // -2
                name = name[idx:]
                m = match(name,name_number,name_date)
                if len(m) > 0:
                    browser.find_element_by_css_selector(sno).send_keys(m[0])
                    browser.find_element_by_css_selector(number).send_keys(buy_number)
                    browser.find_element_by_css_selector(date).send_keys(m[1])
                    # break
                else:
                    browser.find_element_by_css_selector(number).send_keys(buy_number)
                    random_str = ''.join(random.sample(['1','2','3','4','5','6','7','8','9','0','z','y','x','w','v','u','t','s','r','q','p','o','n','m','l','k','j','i','h','g','f','e','d','c','b','a'], 7))
                    date_str = '2022-04-22'
                    browser.find_element_by_css_selector(sno).send_keys(random_str)
                    browser.find_element_by_css_selector(date).send_keys(date_str)
                    #browser.find_element_by_xpath('//*[@id="0"]/td[4]/input').send_keys(buy_number)
                    print("没匹配到!!!",origin_name)
        browser.find_element_by_xpath(x).click()
    for x in collapse:
        browser.find_element_by_xpath(x).click()



    
# td[role='gridcell'] > div > divnth-child(1)
#'+packId+'201906051636113215002
#  > td:nth-child(3) > input
# #\32 * > td.ui-sgcollapsed.sgexpanded > div
# pyinstaller -F -w  input.py

if __name__ == "__main__":
    finish = False
    print("登陆后，请打开“药品配送>订单配送”")
    login()
    input("登陆后请回车")
    while not finish:
        open()
        fill()
        input("配送后请回车")
        

