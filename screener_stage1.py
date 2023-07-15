import time as t
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import requests
import pandas as pd
import lxml

import time
import os

def screener_nbknd():
    print("_____________________________________")
    print("enter company name")
    print("_____________________________________")
    name=str(input())

    #for saving files in desired diretion

    os.chdir(r"D:\major_project\automated_fundamental_analysis\output_data_files\screener")
    os.mkdir(name.upper())
    os.chdir("D:\\major_project\\automated_fundamental_analysis\\output_data_files\\screener\\" + name.upper())


    path ="‪C:\\Program Files (x86)\\chromedriver.exe"
    driver = webdriver.Chrome(path)
    url="https://www.screener.in/"
    driver.get(url)
    driver.maximize_window()
    print(driver.title)
    print("_____________________________________")

    #                           login page button
    search = driver.find_element(By.XPATH,
                                 '/html/body/nav/div[2]/div/div/div/div[2]/div[2]/a[1]')  # this is the location of login button
    # search.send_keys(name) #this one is for searching what you want
    search.send_keys(Keys.ENTER)
    search = driver.find_element(By.XPATH, '/html/body/nav/div[2]/div/div/div/div[2]/div[2]/a[1]')  #this is the location of login button
    #search.send_keys(name) #this one is for searching what you want
    search.send_keys(Keys.ENTER)
    t.sleep(1)
        #print("_____________________________________")
        #                           login page details
    search = driver.find_element(By.XPATH, '//*[@id="id_username"]')
    search.send_keys('aditya009udemy@gmail.com') #adityakorade009@gmail.com
    search = driver.find_element(By.XPATH, '//*[@id="id_password"]')
    search.send_keys('kuro@2247') #kuro@2247)
    search.send_keys(Keys.ENTER)
    t.sleep(1)

    #                           searching page
    search = driver.find_element(By.XPATH, '//*[@id="desktop-search"]/div/input')  #this is the location of login button
    search.send_keys(name) #this one is for searching what you want
    t.sleep(1)
    search.send_keys(Keys.ENTER)
    t.sleep(1)

    #                           expanding tables



    t.sleep(1)
    #                           table scraping

    new_url=driver.current_url #getting the url of the banking statement
    print(driver.current_url)
    response = requests.get(new_url)
    print("execution completed code exited by [0]")
    df = pd.read_html(new_url)
    #print(df)
    t.sleep(1)
    df[0].to_excel(name+'_Quarterly Results.xlsx')
    df[1].to_excel(name+'Profit & Loss.xlsx')
    df[2].to_excel(name+'Compounded Sales Growth.xlsx')
    df[3].to_excel(name+'Compounded Profit Growth.xlsx')
    df[4].to_excel(name+'Stock Price CAGR.xlsx')
    df[5].to_excel(name+'Return on Equity.xlsx')
    df[6].to_excel(name+'Balance Sheet.xlsx')
    df[7].to_excel(name+'Cash Flows.xlsx')
    df[8].to_excel(name+'Cash Flows.xlsx')
    df[9].to_excel(name+'Shareholding Pattern.xlsx')
    #name competitive company
    #//tbody/tr/td[2]/a
    print(driver.find_element(By.XPATH, '//*[@id="top-ratios"]'))
    #getting html code


    driver.close()
    print("execution completed code exited by [0]")
    print("do you want to scrap for another company press 1")
    rep = int(input())
    if (rep==1):
        screener_nbknd()
screener_nbknd()
print("thank you")
