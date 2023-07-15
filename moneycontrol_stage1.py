import time as t
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import requests
import pandas as pd
import lxml.html
import time
from datetime import datetime as dt
import os
import timeit
def moneycontrol_bkend():
    print("_____________________________________")
    print("enter company name")
    print("_____________________________________")
    name=str(input())

    #for saving files in desired diretion

    os.chdir(r"D:\major_project\automated_fundamental_analysis\output_data_files\moneycontrol")
    os.mkdir(name.upper())
    os.chdir("D:\\major_project\\automated_fundamental_analysis\\output_data_files\\moneycontrol\\"+name.upper())

    path = "‪C:\chromedriver\chromedriver.exe"
    driver = webdriver.Chrome(path)
    url = "https://www.moneycontrol.com/"
    driver.get(url)
    driver.maximize_window()
    #t.sleep(1)
    print(driver.title)
    print("_____________________________________")
    #t.sleep(1)
    search = driver.find_element(By.XPATH, '/html/body/div/div[1]/span/a')  # closing ad
    search.send_keys(Keys.ENTER)
    #t.sleep(1)
    search = driver.find_element(By.XPATH, '//*[@id="search_str"]')  # search button
    search.send_keys(name)  # this one is for searching what you want
    search.send_keys(Keys.ENTER)
    #content = driver.page_source
    #print("content variable = "+content)
    #soup = BeautifulSoup(content)
    html=requests.get(driver.current_url)
    print("this is the xpath link")
    doc=lxml.html.fromstring(html.content)
    print("_____________________________________")
    print(doc)
    print("_____________________________________")
    balance_sheet= doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[1]/a/@href')
    profit_and_loss = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[2]/a/@href')
    quarterly_result = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[3]/a/@href')
    half_yearly_result = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[4]/a/@href')
    nine_month_result = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[5]/a/@href')
    yearly_result = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[6]/a/@href')
    cashflow = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[7]/a/@href')
    ratios = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[8]/a/@href')
    capital_structure = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[9]/a/@href')
    df = pd.read_html(balance_sheet[0])
    df[0].to_excel(name + '_Balance_Sheet.xlsx')
    df = pd.read_html(profit_and_loss[0])
    df[0].to_excel(name + '_Profit_And_Loss.xlsx')
    df = pd.read_html(quarterly_result[0])
    df[0].to_excel(name + '_Quarterly_Result.xlsx')
    df = pd.read_html(half_yearly_result[0])
    df[0].to_excel(name + '_Half_Yearly_Result.xlsx')
    df = pd.read_html(nine_month_result[0])
    df[0].to_excel(name + '_Nine_Month_Result.xlsx')
    df = pd.read_html(yearly_result[0])
    df[0].to_excel(name + '_Yearly_result.xlsx')
    df = pd.read_html(cashflow[0])
    df[0].to_excel(name + '_Cashflow.xlsx')
    df = pd.read_html(ratios[0])
    df[0].to_excel(name + '_Ratios.xlsx')
    df = pd.read_html(capital_structure[0])
    df[0].to_excel(name + '_Capital_Structure.xlsx')
    #t.sleep(1)
    driver.close()
    print("execution completed code exited by [0]")
    print("do you want to scrap for another company press 1")
    rep = int(input())
    if (rep == 1):
        moneycontrol_bkend()
moneycontrol_bkend()
print("thank you")

