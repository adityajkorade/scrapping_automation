#add 1 more frame which will get triggered once you click 'open folder 1' it will consist of 9 buttons opening specific location in storage of pc
import tkinter as tk
import time as t
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import requests
import pandas as pd
import lxml.html
import os

def moneycontrol(name):
    print("scraping for ",name,"has started using moneycontrol website")
def screener(name):
    print("scraping for ",name,"has started using screener website")
def news(name):
    print("news scraping for ", name, "has started using screener website")

class App:



    def __init__(self, master):
        self.master = master
        self.master.title("Kurotech_scraper")

        # Create textfield and label
        self.label = tk.Label(self.master, text="Enter Company Symbol:", font=("Arial", 12))
        self.label.pack(pady=10)
        self.entry = tk.Entry(self.master, font=("Arial", 12), width=30)
        self.entry.pack(pady=10)

        # Create checkboxes
        self.var1 = tk.BooleanVar()
        self.var2 = tk.BooleanVar()
        self.var3 = tk.BooleanVar()
        self.check1 = tk.Checkbutton(self.master, text="Money Control", variable=self.var1, font=("Arial", 12))
        self.check2 = tk.Checkbutton(self.master, text="Screener", variable=self.var2, font=("Arial", 12))
        self.check3 = tk.Checkbutton(self.master, text="analyse pre-existing data", variable=self.var3, font=("Arial", 12))
        self.check1.pack(pady=5)
        self.check2.pack(pady=5)
        self.check3.pack(pady=5)

        # Create button to submit choices
        self.button = tk.Button(self.master, text="Submit", font=("Arial", 12), bg="#4CAF50", fg="white",
                                command=self.submit)
        self.button.pack(pady=10)

    def submit(self):
        # Get name from text field
        name = self.entry.get()

        # Create new window to display choices
        self.new_window = tk.Toplevel(self.master)
        self.new_window.title("Choices")
        self.new_window.geometry("500x200")
        self.new_window1 = tk.Toplevel(self.master)
        self.new_window1.title("selected")
        self.new_window1.geometry("500x200")

        # only money control
        if self.var1.get() and not self.var2.get() and not self.var3.get():
            msg = f"{name}, you chose MoneyControl for scraping click scrape button."
            #t.sleep(10)
            label = tk.Label(self.new_window1, text=msg, font=("Arial", 12))
            label.pack(pady=10)

            # Create buttons to initiate scrape
            self.scrape1 = tk.Button(self.new_window1, text="scrape", font=("Arial", 12), bg="#3498DB",
                                fg="white",
                                command=self.scrape1)
            self.ok = tk.Button(self.new_window1, text="ok", font=("Arial", 12), bg="#3498DB",
                                     fg="white",
                                     command=self.quit1)
            self.ok.pack(side=tk.RIGHT, padx=10, pady=10)
            self.scrape1.pack(side=tk.LEFT, padx=10, pady=10)


        #screener
        elif self.var2.get() and not self.var1.get() and not self.var3.get():
            msg = f"{name}, you chose option 2."
            label = tk.Label(self.new_window1, text=msg, font=("Arial", 12))
            label.pack(pady=10)
            # Create buttons to initiate scrape
            self.scrape1 = tk.Button(self.new_window1, text="scrape", font=("Arial", 12), bg="#3498DB",
                                     fg="white",
                                     command=self.scrape2)
            self.ok = tk.Button(self.new_window1, text="ok", font=("Arial", 12), bg="#3498DB",
                                fg="white",
                                command=self.quit1)
            self.ok.pack(side=tk.RIGHT, padx=10, pady=10)
            self.scrape1.pack(side=tk.LEFT, padx=10, pady=10)


        #scrape all
        elif self.var1.get() and self.var2.get() and self.var3.get():
            msg = f"{name}, you chose all the options."
            label = tk.Label(self.new_window1, text=msg, font=("Arial", 12))
            label.pack(pady=10)
            # Create buttons to initiate scrape
            self.scrape1 = tk.Button(self.new_window1, text="scrape", font=("Arial", 12), bg="#3498DB",
                                     fg="white",
                                     command=self.scrape3)
            self.ok = tk.Button(self.new_window1, text="ok", font=("Arial", 12), bg="#3498DB",
                                fg="white",
                                command=self.quit1)
            self.ok.pack(side=tk.RIGHT, padx=10, pady=10)
            self.scrape1.pack(side=tk.LEFT, padx=10, pady=10)


        #scrape analyse news
        elif self.var3.get() and not self.var2.get() and not self.var1.get():
            msg = f"{name}, you chose last option."
            label = tk.Label(self.new_window1, text=msg, font=("Arial", 12))
            label.pack(pady=10)
            # Create buttons to initiate scrape
            self.scrape1 = tk.Button(self.new_window1, text="scrape", font=("Arial", 12), bg="#3498DB",
                                     fg="white",
                                     command=self.scrape4)
            self.ok = tk.Button(self.new_window1, text="ok", font=("Arial", 12), bg="#3498DB",
                                fg="white",
                                command=self.quit1)
            self.ok.pack(side=tk.RIGHT, padx=10, pady=10)
            self.scrape1.pack(side=tk.LEFT, padx=10, pady=10)


        #scrape using analyse and screener
        elif self.var3.get() and self.var2.get() and not self.var1.get():
            msg = f"{name}, you chose second and third options."
            label = tk.Label(self.new_window1, text=msg, font=("Arial", 12))
            label.pack(pady=10)
            # Create buttons to initiate scrape
            self.scrape1 = tk.Button(self.new_window1, text="scrape", font=("Arial", 12), bg="#3498DB",
                                fg="white",
                                command=self.scrape5)
            self.ok = tk.Button(self.new_window1, text="ok", font=("Arial", 12), bg="#3498DB",
                                     fg="white",
                                     command=self.quit1)
            self.ok.pack(side=tk.RIGHT, padx=10, pady=10)
            self.scrape1.pack(side=tk.LEFT, padx=10, pady=10)


        # scrape using analyse and moneycontrol
        elif self.var3.get() and not self.var2.get() and self.var1.get():
            msg = f"{name}, you chose first and third options."
            label = tk.Label(self.new_window1, text=msg, font=("Arial", 12))
            label.pack(pady=10)
            # Create buttons to initiate scrape
            self.scrape1 = tk.Button(self.new_window1, text="scrape", font=("Arial", 12), bg="#3498DB",
                                     fg="white",
                                     command=self.scrape6)
            self.ok = tk.Button(self.new_window1, text="ok", font=("Arial", 12), bg="#3498DB",
                                fg="white",
                                command=self.quit1)
            self.ok.pack(side=tk.RIGHT, padx=10, pady=10)
            self.scrape1.pack(side=tk.LEFT, padx=10, pady=10)


        # scrape using moneycontrol and screener
        elif self.var2.get() and not self.var3.get() and self.var1.get():
            msg = f"{name}, you chose second and first options."
            label = tk.Label(self.new_window1, text=msg, font=("Arial", 12))
            label.pack(pady=10)
            # Create buttons to initiate scrape
            self.scrape1 = tk.Button(self.new_window1, text="scrape", font=("Arial", 12), bg="#3498DB",
                                     fg="white",
                                     command=self.scrape7)
            self.ok = tk.Button(self.new_window1, text="ok", font=("Arial", 12), bg="#3498DB",
                                fg="white",
                                command=self.quit1)
            self.ok.pack(side=tk.RIGHT, padx=10, pady=10)
            self.scrape1.pack(side=tk.LEFT, padx=10, pady=10)


        #just does nothing
        else:
            msg = f", you did not choose any option."
            label = tk.Label(self.new_window1, text=msg, font=("Arial", 12))
            label.pack(pady=10)
            self.ok = tk.Button(self.new_window1, text="ok", font=("Arial", 12), bg="#3498DB",
                                fg="white",
                                command=self.quit1)
            self.ok.pack(side=tk.BOTTOM, padx=10, pady=10)

        label = tk.Label(self.new_window, text="wait for the the scrape window to close", font=("Arial", 12))
        label.pack(pady=10)

        # Create buttons to open folders and update to DB create this one once the new one is updated

        self.button1 = tk.Button(self.new_window, text="open moneycontrol data", font=("Arial", 12), bg="#3498DB", fg="white",
                                 command=self.open_MC)
        self.button2 = tk.Button(self.new_window, text="Open screener data", font=("Arial", 12), bg="#3498DB", fg="white",
                                 command=self.open_SC)
        self.button3 = tk.Button(self.new_window, text="open news data", font=("Arial", 12), bg="#3498DB", fg="white",
                                 command=self.open_news)
        self.button4 = tk.Button(self.new_window, text="Update to DB", font=("Arial", 12), bg="#3498DB", fg="white",
                                 command=self.update_to_db)
        self.button5 = tk.Button(self.new_window, text="reset", font=("Arial", 12), bg="#3498DB", fg="white",
                                 command=self.restart)
        self.button1.pack(side=tk.LEFT, padx=10, pady=10)
        self.button2.pack(side=tk.RIGHT, padx=10, pady=10)
        self.button3.pack(side=tk.TOP, padx=10, pady=10)
        self.button4.pack(side=tk.BOTTOM, padx=10, pady=10)
        self.button5.pack(side=tk.BOTTOM, padx=10, pady=10)











    def scrape1(self):
        # Open specific folder
        name = self.entry.get()
        #moneycontrol(name)
        #print("_____________________________________")
        #print("enter company name")
        ##print("_____________________________________")
        print("starting scraping from the website")
        # for saving files in desired diretion

        os.chdir(r"D:\major_project\automated_fundamental_analysis\output_data_files\moneycontrol")
        os.mkdir(name.upper())
        os.chdir("D:\\major_project\\automated_fundamental_analysis\\output_data_files\\moneycontrol\\" + name.upper())

        path = "‪C:\\Program Files (x86)\\chromedriver.exe"
        driver = webdriver.Chrome(path)
        url = "https://www.moneycontrol.com/"
        driver.get(url)
        driver.maximize_window()
        # t.sleep(1)
        print(driver.title)
        print("_____________________________________")
        # t.sleep(1)
        search = driver.find_element(By.XPATH, '/html/body/div/div[1]/span/a')  # closing ad
        search.send_keys(Keys.ENTER)
        # t.sleep(1)
        search = driver.find_element(By.XPATH, '//*[@id="search_str"]')  # search button
        search.send_keys(name)  # this one is for searching what you want
        search.send_keys(Keys.ENTER)
        # content = driver.page_source
        # print("content variable = "+content)
        # soup = BeautifulSoup(content)
        html = requests.get(driver.current_url)
        print("this is the xpath link")
        doc = lxml.html.fromstring(html.content)
        print("_____________________________________")
        print(doc)
        print("_____________________________________")
        balance_sheet = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[1]/a/@href')
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
        # t.sleep(1)
        driver.close()
        print("execution completed code exited by [0]")
        self.new_window1.destroy()
        print("scraping ended")

    def scrape2(self):
        # Open specific folder
        name = self.entry.get()

        screener(name)
        print("screener scraping started")
        os.chdir(r"D:\major_project\automated_fundamental_analysis\output_data_files\screener")
        os.mkdir(name.upper())
        os.chdir("D:\\major_project\\automated_fundamental_analysis\\output_data_files\\screener\\" + name.upper())

        path = "‪C:\\Program Files (x86)\\chromedriver.exe"
        driver = webdriver.Chrome(path)
        url = "https://www.screener.in/"
        driver.get(url)
        driver.maximize_window()
        print(driver.title)
        print("_____________________________________")

        #                           login page button
        search = driver.find_element(By.XPATH,
                                     '/html/body/nav/div[2]/div/div/div/div[2]/div[2]/a[1]')  # this is the location of login button
        # search.send_keys(name) #this one is for searching what you want
        search.send_keys(Keys.ENTER)
        search = driver.find_element(By.XPATH,
                                     '/html/body/nav/div[2]/div/div/div/div[2]/div[2]/a[1]')  # this is the location of login button
        # search.send_keys(name) #this one is for searching what you want
        search.send_keys(Keys.ENTER)
        t.sleep(1)
        # print("_____________________________________")
        #                           login page details
        search = driver.find_element(By.XPATH, '//*[@id="id_username"]')
        search.send_keys('aditya009udemy@gmail.com')  # adityakorade009@gmail.com
        search = driver.find_element(By.XPATH, '//*[@id="id_password"]')
        search.send_keys('kuro@2247')  # kuro@2247)
        search.send_keys(Keys.ENTER)
        t.sleep(1)

        #                           searching page
        search = driver.find_element(By.XPATH,
                                     '//*[@id="desktop-search"]/div/input')  # this is the location of login button
        search.send_keys(name)  # this one is for searching what you want
        t.sleep(1)
        search.send_keys(Keys.ENTER)
        t.sleep(1)

        #                           expanding tables

        t.sleep(1)
        #                           table scraping

        new_url = driver.current_url  # getting the url of the banking statement
        print(driver.current_url)
        response = requests.get(new_url)
        print("execution completed code exited by [0]")
        df = pd.read_html(new_url)
        # print(df)
        t.sleep(1)
        df[0].to_excel(name + '_Quarterly Results.xlsx')
        df[1].to_excel(name + '_Profit & Loss.xlsx')
        df[2].to_excel(name + '_Compounded Sales Growth.xlsx')
        df[3].to_excel(name + '_Compounded Profit Growth.xlsx')
        df[4].to_excel(name + '_Stock Price CAGR.xlsx')
        df[5].to_excel(name + '_Return on Equity.xlsx')
        df[6].to_excel(name + '_Balance Sheet.xlsx')
        df[7].to_excel(name + '_Cash Flows.xlsx')
        df[8].to_excel(name + '_Cash Flows.xlsx')
        df[9].to_excel(name + '_Shareholding Pattern.xlsx')
        # name competitive company
        # //tbody/tr/td[2]/a
        print(driver.find_element(By.XPATH, '//*[@id="top-ratios"]'))
        # getting html code
        driver.close()
        print("screener scraping completed")
        self.new_window1.destroy()

    def scrape3(self):
        name = self.entry.get()
        # moneycontrol(name)
        # print("_____________________________________")
        # print("enter company name")
        ##print("_____________________________________")
        print("starting scraping from the website")
        # for saving files in desired diretion

        os.chdir(r"D:\major_project\automated_fundamental_analysis\output_data_files\moneycontrol")
        os.mkdir(name.upper())
        os.chdir("D:\\major_project\\automated_fundamental_analysis\\output_data_files\\moneycontrol\\" + name.upper())

        path = "‪C:\\Program Files (x86)\\chromedriver.exe"
        driver = webdriver.Chrome(path)
        url = "https://www.moneycontrol.com/"
        driver.get(url)
        driver.maximize_window()
        # t.sleep(1)
        print(driver.title)
        print("_____________________________________")
        # t.sleep(1)
        search = driver.find_element(By.XPATH, '/html/body/div/div[1]/span/a')  # closing ad
        search.send_keys(Keys.ENTER)
        # t.sleep(1)
        search = driver.find_element(By.XPATH, '//*[@id="search_str"]')  # search button
        search.send_keys(name)  # this one is for searching what you want
        search.send_keys(Keys.ENTER)
        # content = driver.page_source
        # print("content variable = "+content)
        # soup = BeautifulSoup(content)
        html = requests.get(driver.current_url)
        print("this is the xpath link")
        doc = lxml.html.fromstring(html.content)
        print("_____________________________________")
        print(doc)
        print("_____________________________________")
        balance_sheet = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[1]/a/@href')
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
        strenght = doc.xpath('//*[@id="swot_ls"]/a/strong/text()')
        weakness = doc.xpath('//*[@id="swot_lw"]/a/strong/text()')
        opportunity = doc.xpath('//*[@id="swot_lo"]/a/strong/text()')
        threat = doc.xpath('//*[@id="swot_lt"]/a/strong/text()')
        # t.sleep(1)
        driver.close()
        print("execution completed code exited by [0]")
        print("money control scraping ended")


        #-------------------------------screener
        screener(name)

        print("screener scraping started")
        os.chdir(r"D:\major_project\automated_fundamental_analysis\output_data_files\screener")
        os.mkdir(name.upper())
        os.chdir("D:\\major_project\\automated_fundamental_analysis\\output_data_files\\screener\\" + name.upper())

        path = "‪C:\\Program Files (x86)\\chromedriver.exe"
        driver = webdriver.Chrome(path)
        url = "https://www.screener.in/"
        driver.get(url)
        driver.maximize_window()
        print(driver.title)
        print("_____________________________________")

        #                           login page button
        search = driver.find_element(By.XPATH,
                                     '/html/body/nav/div[2]/div/div/div/div[2]/div[2]/a[1]')  # this is the location of login button
        # search.send_keys(name) #this one is for searching what you want
        search.send_keys(Keys.ENTER)
        search = driver.find_element(By.XPATH,
                                     '/html/body/nav/div[2]/div/div/div/div[2]/div[2]/a[1]')  # this is the location of login button
        # search.send_keys(name) #this one is for searching what you want
        search.send_keys(Keys.ENTER)
        t.sleep(1)
        # print("_____________________________________")
        #                           login page details
        search = driver.find_element(By.XPATH, '//*[@id="id_username"]')
        search.send_keys('aditya009udemy@gmail.com')  # adityakorade009@gmail.com
        search = driver.find_element(By.XPATH, '//*[@id="id_password"]')
        search.send_keys('kuro@2247')  # kuro@2247)
        search.send_keys(Keys.ENTER)
        t.sleep(1)

        #                           searching page
        search = driver.find_element(By.XPATH,
                                     '//*[@id="desktop-search"]/div/input')  # this is the location of login button
        search.send_keys(name)  # this one is for searching what you want
        t.sleep(1)
        search.send_keys(Keys.ENTER)
        t.sleep(1)

        #                           expanding tables

        t.sleep(1)
        #                           table scraping

        new_url = driver.current_url  # getting the url of the banking statement
        print(driver.current_url)
        response = requests.get(new_url)
        print("execution completed code exited by [0]")
        df = pd.read_html(new_url)
        # print(df)
        t.sleep(1)
        df[0].to_excel(name + '_Quarterly Results.xlsx')
        df[1].to_excel(name + '_Profit & Loss.xlsx')
        df[2].to_excel(name + '_Compounded Sales Growth.xlsx')
        df[3].to_excel(name + '_Compounded Profit Growth.xlsx')
        df[4].to_excel(name + '_Stock Price CAGR.xlsx')
        df[5].to_excel(name + '_Return on Equity.xlsx')
        df[6].to_excel(name + '_Balance Sheet.xlsx')
        df[7].to_excel(name + '_Cash Flows.xlsx')
        df[8].to_excel(name + '_Cash Flows.xlsx')
        df[9].to_excel(name + '_Shareholding Pattern.xlsx')
        # name competitive company
        # //tbody/tr/td[2]/a
        print(driver.find_element(By.XPATH, '//*[@id="top-ratios"]'))
        # getting html code
        driver.close()
        print("screener scraping completed")
        self.new_window1.destroy()

        #---------------------swot
        print(strenght)
        print(weakness)
        print(opportunity)
        print(threat)

    def scrape4(self):
        # Open specific folder
        news(name = self.entry.get())
        self.new_window1.destroy()

    def scrape5(self):
        # Open specific folder
        news(name = self.entry.get())
        screener(name=self.entry.get())
        self.new_window1.destroy()

    def scrape6(self):
        # Open specific folder
        moneycontrol(name = self.entry.get())
        news(name=self.entry.get())
        self.new_window1.destroy()

    def scrape7(self):
        name = self.entry.get()
        # moneycontrol(name)
        # print("_____________________________________")
        # print("enter company name")
        ##print("_____________________________________")
        print("starting scraping from the website")
        # for saving files in desired diretion

        os.chdir(r"D:\major_project\automated_fundamental_analysis\output_data_files\moneycontrol")
        os.mkdir(name.upper())
        os.chdir("D:\\major_project\\automated_fundamental_analysis\\output_data_files\\moneycontrol\\" + name.upper())

        path = "‪C:\\Program Files (x86)\\chromedriver.exe"
        driver = webdriver.Chrome(path)
        url = "https://www.moneycontrol.com/"
        driver.get(url)
        driver.maximize_window()
        # t.sleep(1)
        print(driver.title)
        print("_____________________________________")
        # t.sleep(1)
        search = driver.find_element(By.XPATH, '/html/body/div/div[1]/span/a')  # closing ad
        search.send_keys(Keys.ENTER)
        # t.sleep(1)
        search = driver.find_element(By.XPATH, '//*[@id="search_str"]')  # search button
        search.send_keys(name)  # this one is for searching what you want
        search.send_keys(Keys.ENTER)
        # content = driver.page_source
        # print("content variable = "+content)
        # soup = BeautifulSoup(content)
        html = requests.get(driver.current_url)
        print("this is the xpath link")
        doc = lxml.html.fromstring(html.content)
        print("_____________________________________")
        print(doc)
        print("_____________________________________")
        balance_sheet = doc.xpath('//*[@id="consolidated"]/div[2]/div[2]/div/ul/li[1]/a/@href')
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
        # t.sleep(1)
        driver.close()
        print("execution completed code exited by [0]")

        print("scraping ended")
        screener(name)

        print("screener scraping started")
        os.chdir(r"D:\major_project\automated_fundamental_analysis\output_data_files\screener")
        os.mkdir(name.upper())
        os.chdir("D:\\major_project\\automated_fundamental_analysis\\output_data_files\\screener\\" + name.upper())

        path = "‪C:\\Program Files (x86)\\chromedriver.exe"
        driver = webdriver.Chrome(path)
        url = "https://www.screener.in/"
        driver.get(url)
        driver.maximize_window()
        print(driver.title)
        print("_____________________________________")

        #                           login page button
        search = driver.find_element(By.XPATH,
                                     '/html/body/nav/div[2]/div/div/div/div[2]/div[2]/a[1]')  # this is the location of login button
        # search.send_keys(name) #this one is for searching what you want
        search.send_keys(Keys.ENTER)
        search = driver.find_element(By.XPATH,
                                     '/html/body/nav/div[2]/div/div/div/div[2]/div[2]/a[1]')  # this is the location of login button
        # search.send_keys(name) #this one is for searching what you want
        search.send_keys(Keys.ENTER)
        t.sleep(1)
        # print("_____________________________________")
        #                           login page details
        search = driver.find_element(By.XPATH, '//*[@id="id_username"]')
        search.send_keys('aditya009udemy@gmail.com')  # adityakorade009@gmail.com
        search = driver.find_element(By.XPATH, '//*[@id="id_password"]')
        search.send_keys('kuro@2247')  # kuro@2247)
        search.send_keys(Keys.ENTER)
        t.sleep(1)

        #                           searching page
        search = driver.find_element(By.XPATH,
                                     '//*[@id="desktop-search"]/div/input')  # this is the location of login button
        search.send_keys(name)  # this one is for searching what you want
        t.sleep(1)
        search.send_keys(Keys.ENTER)
        t.sleep(1)

        #                           expanding tables

        t.sleep(1)
        #                           table scraping

        new_url = driver.current_url  # getting the url of the banking statement
        print(driver.current_url)
        response = requests.get(new_url)
        print("execution completed code exited by [0]")
        df = pd.read_html(new_url)
        # print(df)
        t.sleep(1)
        df[0].to_excel(name + '_Quarterly Results.xlsx')
        df[1].to_excel(name + '_Profit & Loss.xlsx')
        df[2].to_excel(name + '_Compounded Sales Growth.xlsx')
        df[3].to_excel(name + '_Compounded Profit Growth.xlsx')
        df[4].to_excel(name + '_Stock Price CAGR.xlsx')
        df[5].to_excel(name + '_Return on Equity.xlsx')
        df[6].to_excel(name + '_Balance Sheet.xlsx')
        df[7].to_excel(name + '_Cash Flows.xlsx')
        df[8].to_excel(name + '_Cash Flows.xlsx')
        df[9].to_excel\
            (name + '_Shareholding Pattern.xlsx')
        # name competitive company
        # //tbody/tr/td[2]/a
        print(driver.find_element(By.XPATH, '//*[@id="top-ratios"]'))
        # getting html code
        driver.close()
        print("screener scraping completed")
        self.new_window1.destroy()

    def open_MC(self):
        # Open specific folder
        name = self.entry.get().upper()
        path_mc="D:\\major_project\\automated_fundamental_analysis\\output_data_files\\moneycontrol\\"+name
        path_mc=os.path.realpath(path_mc)
        os.startfile(path_mc)

    def open_SC(self):
        name = self.entry.get().upper()
        path_sc = "D:\\major_project\\automated_fundamental_analysis\\output_data_files\\screener\\" + name
        path_sc = os.path.realpath(path_sc)
        os.startfile(path_sc)

    def quit(self):
        # Close both windows and end program
        self.master.destroy()

    def quit1(self):
        # Close both windows and end program
        self.new_window1.destroy()

    def update_to_db(self):
        # Write code to update choices to database
        pass

    def open_news(self):
        # Open specific folder
        import os
        os.system('start "" "C:/path/to/folder1"')

    def restart(self):
        # Close choices window and reset text field and checkboxes
        self.new_window.destroy()
        self.entry.delete(0, tk.END)
        self.var1.set(False)
        self.var2.set(False)
        self.var3.set(False)




# Create main window and run app
root = tk.Tk()
root.geometry("400x300")
app = App(root)
root.mainloop()

