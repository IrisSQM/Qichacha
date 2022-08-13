#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企查查自动化-养老

Created on Wed Jul 27 12:35:45 2022

@author: shiqimeng
"""
# 装包
import os, csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
from random import uniform

# Save path and output file names
save_path = '<path>'

products = save_path + os.sep + '品牌产品库.csv'
finance = save_path + os.sep + '融资信息库.csv'

# Write column heads into csv files
with open(products, 'w', newline = '') as output:
    writer = csv.writer(output)
    writer.writerow(['品牌产品','行业标签','融资信息','融资次数','成立日期','所属地',
                     '所属企业','主营业务','业务简介','官网','企查查公司页','qcc三级行业'])
    
with open(finance, 'w', newline = '') as output:
    writer = csv.writer(output)
    writer.writerow(['品牌产品','融资序号','日期','融资轮次','估值金额','融资金额',
                     '投资机构'])
    
# ----------------- Scrape Products Page 定义函数 -------------------------------------
# first go to products' home page
def hm_login(user, pw):
    '''
    自动打开网页，输入用户名、密码
    **手动**验证
    
    Input: user, pw

    Returns
    -------
    None.

    '''
    hm_path = 'https://www.qcc.com/web/project/invest-org/application/classify?tab=2'
    browser.get(hm_path)
    time.sleep(uniform(2,2.3))
    
    # log in
    # switch to password login
    other_login = browser.find_element(by=By.XPATH, 
                value='/html/body/div[1]/div[2]/div[2]/div/div[3]/img')
    other_login.click()
    
    other_login_pw = browser.find_element(by=By.XPATH, 
                value='/html/body/div[1]/div[2]/div[2]/div/div[1]/div[1]/div[2]/a')
    other_login_pw.click()
    
    # fill in user & key
    browser.find_element(by=By.XPATH, 
                value='/html/body/div[1]/div[2]/div[2]/div/div[1]/div[3]/form/div[1]/input'
                ).send_keys(user)
    
    browser.find_element(by=By.XPATH, 
                value='/html/body/div[1]/div[2]/div[2]/div/div[1]/div[3]/form/div[2]/input'
                ).send_keys(pw)
    
    browser.find_element(by=By.XPATH, 
                value='/html/body/div[1]/div[2]/div[2]/div/div[1]/div[3]/form/div[4]/button'
                ).click()
    
    time.sleep(uniform(20,25)) # 利用这段时间完成人工验证

# 筛选条件
def filter12():
    '''
    筛选一级、二级行业
    一级：医疗
    二级：养老

    Returns
    -------
    None.

    '''
    # 所在一级行业
    browser.find_element(by=By.XPATH, 
                value='/html/body/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div[4]/div[2]/div/div/div[2]/span[4]/a'
                ).click() 
    browser.execute_script('window.scrollTo(1000,0)')
    # 二级行业
    # 先展开更多
    browser.find_element(by=By.XPATH, 
                value='/html/body/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div[3]/div[2]/div/div[2]/div[3]/a/span'
                ).click()
    # 再选择养老
    browser.find_element(by=By.XPATH, 
                value='/html/body/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div[3]/div[2]/div/div[2]/div[2]/span[17]/a'
                ).click()

# 因为查询最多100页，所以需要按照三级行业爬取
def filter3(i):
    '''
    筛选三级行业
    Parameters
    ----------
    i : Int
        三级行业序号，从 1 开始

    Returns
    -------
    None.

    '''
    f3 = browser.find_element(by=By.XPATH, 
                value='/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div[1]/div/div[3]/div[2]/div/div[3]/div[2]/span[{}]/a'.format(str(i))
                )
    time.sleep(uniform(0.5,0.8))
    f3.click()

# ----------------- -------------------------------------
# 依次进入产品页爬取内容
def scrapeprod(row_no):
    '''
    Scrape a product

    Parameters
    ----------
    row_no : int from 1

    Returns
    -------
    None.

    '''
    
    # 记录当前页面
    main_win = browser.current_window_handle
    
    # 进入产品页
    browser.find_element(by=By.XPATH, 
                value='/html/body/div/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/table/tr[{}]/td[2]/a/div/div[1]/span'.format(str(row_no))
                ).click()
    
    browser.switch_to.window(browser.window_handles[1])
    
    time.sleep(uniform(1,2))
    
    # 抓取信息
    try:
        prod_name = browser.find_element(by=By.XPATH,
                    value='/html/body/div/div[2]/section/div[2]/h1/span'
                    ).text
    except:
        time.sleep(uniform(20,30))  # 手动验证
        prod_name = browser.find_element(by=By.XPATH,
                    value='/html/body/div/div[2]/section/div[2]/h1/span'
                    ).text
    
    try:
        comp_web = browser.find_element(by=By.XPATH,
                    value='/html/body/div/div[2]/section/div[2]/h1/a'
                    ).get_attribute('href')
    except:
        comp_web = '-'
    
    ind_tags_re = browser.find_elements(by=By.XPATH,
                value='/html/body/div/div[2]/section/div[2]/div/div[1]'
                )
    ind_tags = ' '.join([tag.text for tag in ind_tags_re])
    
    fin_info = browser.find_element(by=By.XPATH,
                value='/html/body/div/div[2]/section/div[2]/div/div[2]/span[1]/span'
                ).text
    
    found_dt = browser.find_element(by=By.XPATH,
                value='/html/body/div/div[2]/section/div[2]/div/div[2]/span[2]/span'
                ).text
    
    location = browser.find_element(by=By.XPATH,
                value='/html/body/div/div[2]/section/div[2]/div/div[2]/span[3]/span'
                ).text
    
    try:
        comp = browser.find_element(by=By.XPATH,
                    value='/html/body/div/div[2]/section/div[2]/div/div[3]/span/span/a'
                    ).text
    except:
        comp = '-'
    
    try:
        main_busi = browser.find_element(by=By.XPATH,
                    value='/html/body/div/div[2]/section/div[2]/div/div[4]/span'
                    ).text
    except:
        main_busi = '-'
    
    busi_desc = browser.find_element(by=By.XPATH,
                value='/html/body/div/div[2]/div[1]/div[1]/section/section[1]/table/tbody/tr/td'
                ).text
    
    try:
        qcc_web = browser.find_element(by=By.XPATH,
                    value='/html/body/div/div[2]/section/div[2]/div/div[3]/span/span/a'
                    ).get_attribute('href')
    except:
        qcc_web = '-'
    
    if fin_info == '-':
        fin_no = '0'
    
    else:
        fin_no = browser.find_element(by=By.XPATH,
                    value='/html/body/div/div[2]/div[1]/div[1]/section/section[2]/div[1]/span[1]'
                    ).text
    
        for fin_row_no in range(int(fin_no)):
            row_list_re = browser.find_elements(by=By.XPATH,
                        value='/html/body/div/div[2]/div[1]/div[1]/section/section[2]/div[2]/table/tr[{}]'
                        .format(fin_row_no + 2)
                        )
            row_list_str = row_list_re[0].text
            row_list = row_list_str.split(' ',5)[:-1]+row_list_str.split(' ',5)[-1].rsplit(' ',1)[0:1]
            
            # 融资历史写入csv
            with open(finance, 'a', newline = '') as output:
                 writer = csv.writer(output)
                 writer.writerow([prod_name]+row_list)
    
    # 品牌写入csv
    with open(products, 'a', newline = '') as output:
         writer = csv.writer(output)
         writer.writerow([prod_name,ind_tags,fin_info,fin_no,found_dt,location,
                          comp,main_busi,busi_desc,comp_web,qcc_web,f3])
    
    # 关闭产品页
    browser.close()
    
    # 回到主界面
    browser.switch_to.window(main_win)
    
    # 爬取下一个产品,中间间隔 3-4 秒不等
    time.sleep(uniform(3,4))

# Scrape products under a certain 3rd level filter
def scrapefilter(f3):
    '''
    Scrape products under a certain 3rd level filter

    Parameters
    ----------
    f3 : int from 1
        filter number

    Returns
    -------
    None.

    '''
    
    filter3(f3)
    time.sleep(uniform(0.8,1))
    
    total_prod = int(browser.find_element(by=By.XPATH,
                 value='/html/body/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/h4/span'
                 ).text)
    
    last_page_rows = total_prod % 10
    click_no = total_prod // 10
    
    # scrape first page
    for row_no in range(1,11):
        scrapeprod(row_no)
    
    print('Finish scraping page 1...')
    
    next_page = browser.find_element(by=By.XPATH,
                value='/html/body/div/div[2]/div[2]/div/div/div[2]/ul/li[7]/a')
    next_page.click()
    time.sleep(uniform(0.5,0.8))
    
    for i in range(click_no-1):
        if i%2 == 0:
            time.sleep(uniform(20,30))
        
        for row_no in range(1,11):
            scrapeprod(row_no)
        
        print('Finish scraping page {}...'.format(i+2))
        
        next_page = browser.find_element(by=By.XPATH,
                     value='/html/body/div/div[2]/div[2]/div/div/div[2]/ul/li[8]/a')
        next_page.click()
        
        time.sleep(uniform(0.5,0.8))
    
    # scrape last page
    for row_no in range(1,last_page_rows+1):
        scrapeprod(row_no)

# ----------------- Scrape Products Page 主程序 -------------------------------------
# Set webdriver
driver_path = Service('<driver path>') #输入driver path
browser = webdriver.Chrome(service = driver_path)

# 定义用户名、密码
user = '<用户名>'
pw = '<密码>'

hm_login(user,pw)
filter12()

for f3 in range(1,5):
    print('****** Start scraping filter {} ******'.format(f3))
    scrapefilter(f3)
    print('****** Finish scraping filter {} ******'.format(f3))

