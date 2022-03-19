from selenium import webdriver
from webdriver_manager import driver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
import string

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement

import unittest
from openpyxl import load_workbook

options = webdriver.ChromeOptions()

options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option("excludeSwitches",["enable-automation"])


driver_path = '.\chromedriver_win32\chromedriver.exe'
web = webdriver.Chrome(executable_path=driver_path, chrome_options=options)

def start():
    # Add your credentials here
    username = ""
    password = ""

    web.get('http://de.ingrammicro.int/site/Login/Login?ReturnUrl=%2fSite%2f')

    select = Select(web.find_element_by_id('LanguageDdl'))

    time.sleep(2)
    
    select.select_by_value('de-DE')

    time.sleep(1)

    user = web.find_element_by_id('UserName').send_keys(username)
    passwrd = web.find_element_by_id('Password').send_keys(password)

    login = web.find_element_by_xpath('/html/body/div[1]/div/section[2]/div/div[2]/form/div[3]/div[2]/button').click()

    delay = 99
    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/section[2]/h2')))

    except TimeoutException:
        print ("Loading took too much time!")

# ------------------------------------------------------

def close_chrome():
    web.quit()
    print("Chrome closed.")
# ------------------------------------------------------


def central_banner(page_url, imgName, strDate, endDate, x):

    web.get(page_url)
    time.sleep(3)

    print("Line ", x, ": ", imgName)

    name = web.find_element_by_id("Name").clear()
    name= web.find_element_by_id("Name").send_keys(imgName)
    
    start_date = web.find_element_by_id("StartDate").send_keys(Keys.CONTROL, 'a')
    time.sleep(1)
    start_date = web.find_element_by_id("StartDate").send_keys(strDate)

    end_date = web.find_element_by_id("EndDate").send_keys(Keys.CONTROL, 'a')
    time.sleep(1)
    end_date = web.find_element_by_id("EndDate").send_keys(endDate)
    
    time.sleep(2)

    save_btn = web.find_element_by_id("btnSubmit").click()

    delay= 60
    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/section[2]/div[4]/div[1]/h3')))
        print("Bannet set up \n____________________________________________________")

    except TimeoutException:
        print ("Loading took too much time!")
                                       

# -----------------------------------------------

def middle_banner_search(page_url, imgName, strDate, endDate, x):

    web.get(page_url)
    time.sleep(3)

    print("Line ", x, ": ", imgName)

    name = web.find_element_by_id("Name").clear()
    name= web.find_element_by_id("Name").send_keys(imgName)


    start_date = web.find_element_by_id("StartDate").send_keys(Keys.CONTROL, 'a')
    time.sleep(1)
    start_date = web.find_element_by_id("StartDate").send_keys(strDate)

    end_date = web.find_element_by_id("EndDate").send_keys(Keys.CONTROL, 'a')
    time.sleep(1)
    end_date = web.find_element_by_id("EndDate").send_keys(endDate)
    
    time.sleep(2)

    save_btn = web.find_element_by_id("btnSubmit").click()

    delay= 60
    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/section[2]/div[4]/div[1]/h3')))
        print("Bannet set up \n____________________________________________________")

    except TimeoutException:
        print ("Loading took too much time!")

# -----------------------------------------------


def middle_banner_details(page_url, imgName, strDate, endDate, x):

    web.get(page_url)
    time.sleep(3)

    print("Line ", x, ": ", imgName)

    name = web.find_element_by_id("Name").clear()
    name= web.find_element_by_id("Name").send_keys(imgName)


    placement_zone = Select(web.find_element_by_id('PlacementZone'))
    time.sleep(1)
    placement_zone.select_by_value('0')

    start_date = web.find_element_by_id("StartDate").send_keys(Keys.CONTROL, 'a')
    time.sleep(1)
    start_date = web.find_element_by_id("StartDate").send_keys(strDate)

    end_date = web.find_element_by_id("EndDate").send_keys(Keys.CONTROL, 'a')
    time.sleep(1)
    end_date = web.find_element_by_id("EndDate").send_keys(endDate)
    
    time.sleep(2)

    save_btn = web.find_element_by_id("btnSubmit").click()

    delay= 50
    try:                                                                              
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/section[2]/div[4]/div[1]/h3')))
        print("Bannet set up \n____________________________________________________")

    except TimeoutException:
        print ("Loading took too much time!")

# -----------------------------------------------
def skyscrapper(page_url, imgName, strDate, endDate, x):

    web.get(page_url)
    time.sleep(3)

    print("Line ", x, ": ", imgName)

    name = web.find_element_by_id("Name").clear()
    name= web.find_element_by_id("Name").send_keys(imgName)
    
    start_date = web.find_element_by_id("StartDate").send_keys(Keys.CONTROL, 'a')
    time.sleep(1)
    start_date = web.find_element_by_id("StartDate").send_keys(strDate)

    end_date = web.find_element_by_id("EndDate").send_keys(Keys.CONTROL, 'a')
    time.sleep(1)
    end_date = web.find_element_by_id("EndDate").send_keys(endDate)
    
    time.sleep(2)

    save_btn = web.find_element_by_id("btnSubmit").click()

    delay= 30
    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/section[2]/div[4]/div[1]/h3')))
        print("Bannet set up \n____________________________________________________")

    except TimeoutException:
        print ("Loading took too much time!")

def final ():
    web.get('https://great-success.netlify.app')