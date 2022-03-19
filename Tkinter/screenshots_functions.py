from selenium import webdriver
from webdriver_manager import driver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
import string
import pyautogui
import sys

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

def get_info():
    return "Check Mate"

def screenshot (imgName):
    image_loc = r'.\Screenshots' + r"\\" + imgName
    pyautogui.screenshot(imageFilename=image_loc)

    print("Screenshot made", "\n______________________________________________", flush=True)


def cookies_eu ():
    username = web.find_element_by_xpath('//*[@id="onetrust-accept-btn-handler"]').click()
    web.maximize_window()

def cookies_com():
    username = web.find_element_by_xpath('//*[@id="IMOnlineMvc_V2"]/body/div[1]/div/a').click()
    web.maximize_window()

def work_on(x, imgName):
    print("Line ", x, ": ", imgName)
# ------------------------------------------------------

def sponsored_news(page_url, imgName, x):

    web.get(page_url)
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
   
    delay = 8 # seconds

    
    try:
        cookies_eu ()
    except:
        pass

    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'news-container')))
    except TimeoutException:
        print ("Loading took too much time!")
    
    element = web.find_element_by_id('news-container')
    # element = web.find_element_by_class_name('text-muted')   

    desired_y = (element.size['height'] / 2) + element.location['y']
    current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    scroll_y_by = desired_y - current_y
    web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 

    

    # Take the screenshot -----------------------------
    time.sleep(2)
    
    screenshot(imgName)

    return ('done')

def sponsored_events(page_url, imgName, x):

    web.get(page_url)
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
    try:
        cookies_eu()
    except:
        pass

    delay = 8 # seconds

    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'news-container')))
    except TimeoutException:
        print ("Loading took too much time!")


    element = web.find_element_by_id('bootstrapcolumnlayout_org')

    desired_y = (element.size['height'] / 2) + element.location['y']
    current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    scroll_y_by = desired_y - current_y
    web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 

    # Take the screenshot -----------------------------
    time.sleep(6)
    
    screenshot(imgName)

    return ('done')

def spotlight_home_banner(page_url, imgName, x):

    web.get(page_url)
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
    delay = 10
    
    try:
        cookies_com()
    except:
        pass
                                            
    element = web.find_element_by_id('divHomePageAdsMiddlePreSpotLight')

    desired_y = (element.size['height'] / 2) + element.location['y']
    current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    scroll_y_by = desired_y - current_y
    web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 
    

    # Take the screenshot -----------------------------
    time.sleep(6)
    
    screenshot(imgName)

    return ('done')

def top_home(page_url, imgName, x):

    web.get(page_url)
    time.sleep(4)

    work_on(x, imgName)
    # Accept cookies ------------------------------
    delay= 10
    try:
        cookies_com()
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'p_lt_ctl00_pageplaceholder_p_lt_ctl03_BannerTopSlim_lnkBanner')))
    except:
        pass
                             
    web.execute_script("window.scrollTo(0, 0)")  
    

    # Take the screenshot -----------------------------
    time.sleep(3)
    
    screenshot(imgName)

    return ('done')

def xl_slider_banner_1(page_url, imgName, x):

    web.get(page_url)
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
   
    try:
        cookies_com()
    except:
        pass
                                            
    web.execute_script("window.scrollTo(0, 0)")

    # Take the screenshot -----------------------------
    time.sleep(2)
    
    screenshot(imgName)

    return ('done')

def xl_slider_banner_2(page_url, imgName, x):

    web.get(page_url)
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
   
    
    try:
        cookies_com()
    except:
        pass
                                            

    web.execute_script("window.scrollTo(0, 0)")

    # Take the screenshot -----------------------------
    time.sleep(8)
    
    screenshot(imgName)

    return ('done')

def xl_slider_banner_3(page_url, imgName, x):

    web.get(page_url)
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
   
    
    try:
        cookies_com()
    except:
        pass
                                            
    web.execute_script("window.scrollTo(0, 0)")

    # Take the screenshot -----------------------------
    time.sleep(16)
    
    screenshot(imgName)

    return ('done')
# -----------------------------------------------
def uberall_banner_1(imgName, x): #search zone

    web.get('https://de.ingrammicro.com/site/Search#')
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
   
    delay = 8 # seconds

    try:                                        
        cookies_com()
    except:
        pass
    
    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'searchpagezone-bottombanner')))
    except TimeoutException:
        print ("Loading took too much time!")
        uberall_banner_1(imgName, x)

    element = web.find_element_by_id('searchpagezone-bottombanner')
  
    desired_y = (element.size['height'] / 2) + element.location['y']
    current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    scroll_y_by = desired_y - current_y
    web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 

    # Take the screenshot -----------------------------
    time.sleep(2)
    
    screenshot(imgName)

    return ('done')

def uberall_banner_2(imgName, x): #product zone

    web.get('https://de.ingrammicro.com/site/Search#')
    
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
   
    delay = 8 # seconds

    try:
        cookies_com()
    except:
        pass
    time.sleep(3)

    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="searchResults"]/div[4]/div[1]/div/div')))
    except TimeoutException:
        print ("Loading took too much time!")
        uberall_banner_2(imgName, x)

    product = web.find_element_by_xpath('//*[@id="searchResults"]/div[4]/div[1]/div/div').click()

    time.sleep(3)

    try:
        myEl = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'BottomBanner')))
    except TimeoutException:
        print ("Loading Product details banner took too much time!")
        uberall_banner_2(imgName, x)

    element = web.find_element_by_id('BottomBanner')

    web.execute_script("arguments[0].scrollIntoView();", element) 
    
    desired_y = (element.size['height'] / 2) + element.location['y']
    current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    scroll_y_by = desired_y - current_y
    web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 

    # Take the screenshot -----------------------------
    time.sleep(2)
    
    screenshot(imgName)

    return ('done')
def uberall_banner_3(imgName, x): #home

    web.get('https://de.ingrammicro.com/Site/home')
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
   
    delay = 8 # seconds

    try:
        cookies_com()
    except:
        pass
    
    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'bottomAdBanner')))
    except TimeoutException:
        print ("Loading took too much time!")
        uberall_banner_3(imgName, x)

    element = web.find_element_by_id('bottomAdBanner')
  
    desired_y = (element.size['height'] / 2) + element.location['y']
    current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    scroll_y_by = desired_y - current_y
    web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 

    # Take the screenshot -----------------------------
    time.sleep(2)
    
    screenshot(imgName)

    return ('done')
# -----------------------------------------------
def top_banner_paket(page_url, imgName, x):

    web.get(page_url)
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
   
    
    try:
        cookies_com()
    except:
        pass

    time.sleep(2)

    web.execute_script("window.scrollTo(0, 0)")

    # Take the screenshot -----------------------------
    time.sleep(3)
    
    screenshot(imgName)

    return ('done')

def top_banner_paket_1(page_url, imgName, x):

    web.get(page_url)
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
   
    
    try:
        cookies_com()
    except:
        pass

    delay= 8

    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'topAdBanner')))

    except TimeoutException:
        print ("Loading took too much time!")
        top_banner_paket_1(page_url, imgName, x)

    time.sleep(2)    

    username = web.find_element_by_xpath('//*[@id="searchResults"]/div[4]/div[1]/div/div ').click() 

    time.sleep(2) 

    web.execute_script("window.scrollTo(0, 0)")

    # Take the screenshot -----------------------------
    time.sleep(3)
    
    screenshot(imgName)
    return ('done')

def central_banner(page_url, imgName, x):

    web.get(page_url)
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
   
    
    try:
        cookies_com()
    except:
        pass
                                            
    web.execute_script("window.scrollTo(0, 0)")

    # Take the screenshot -----------------------------
    time.sleep(5)
    
    screenshot(imgName)

    return ('done')
# -----------------------------------------------
def spotlight_banner_premium_1(page_url, imgName, x):

    web.get(page_url)
    
    time.sleep(5)

    work_on(x, imgName)
    # Accept cookies ------------------------------
    try:
        cookies_com()
    except:
        pass

    delay = 10
    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'MiddleBanner')))
        secondElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'searchpagezone-bottombanner')))

    except TimeoutException:
        print ("Loading took too much time!")
        spotlight_banner_premium_1(page_url, imgName, x)

    time.sleep(4)

    element = web.find_element_by_xpath('//*[@id="searchpagezone-middlebanner"]')

    # --------------------------------------------------------------
    # web.execute_script("arguments[0].scrollIntoView();", element) 
    # web.execute_script("window.scrollBy(0, -100)")
    # --------------------------------------------------------------

    # desired_y = (element.size['height'] / 2) + element.location['y']
    # current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    # scroll_y_by = desired_y - current_y
    # web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 
    # web.execute_script("window.scrollBy(0, -100)")

    # ----------------------------------------------------------------------
    actions = ActionChains(web)
    actions.move_to_element(element).perform()

    # web.execute_script("window.scrollBy(0, -100)")

    time.sleep(2)

    desired_y = (element.size['height'] / 2) + element.location['y']
    current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    scroll_y_by = desired_y - current_y + 100  # това +100 е под въпрос
    web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 


    # web.execute_script("window.scrollBy(0, -100)")

    # Take the screenshot -----------------------------
    time.sleep(4)
    
    screenshot(imgName)

    return ('done')

def spotlight_banner_details(page_url, imgName, x):

    web.get(page_url)
    time.sleep(2)

    work_on(x, imgName)
    # Accept cookies ------------------------------
    try:
        cookies_com()
    except:
        pass
    
    delay = 8
    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="searchResults"]/div[4]/div[1]/div/div')))
    except TimeoutException:
        print ("Loading took too much time!")
    
    product = web.find_element_by_xpath('//*[@id="searchResults"]/div[4]/div[1]/div/div/div[1]/a').click()

    time.sleep(3)

    # Check screen resolution
    import ctypes
    user32 = ctypes.windll.user32
    screensizeWidth = user32.GetSystemMetrics(0)
    
    if screensizeWidth <= 1920: 

        web.execute_script("document.body.style.zoom='80%'") 
    #-----------------------------------------
    
    time.sleep(2)

    delay = 10
    try:
        myEl = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'productdetailspagezone-middlebanner')))
    except TimeoutException:
        spotlight_banner_details(page_url, imgName, x)
        print ("Loading took too much time!")
                                            
    # element = web.find_element_by_id('SearchPageZone-MiddleBanner')

    element = web.find_element_by_xpath('//*[@id="productDetailMainSection"]/div[1]/div[2]')

    web.execute_script("arguments[0].scrollIntoView();", element) 
    

    # Take the screenshot -----------------------------
    time.sleep(6)
    
    screenshot(imgName)

    return ('done')

def skyscrapper_A(page_url, imgName, x):

    web.get(page_url)
    time.sleep(3)

    delay = 8 # seconds

    work_on(x, imgName)
    # Accept cookies ------------------------------
    try:
        cookies_com()
    except:
        pass

    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'AllPriceClear')))
    except TimeoutException:
        print ("Loading took too much time!")
                                        
    element = web.find_element_by_xpath('//*[@id="searchMainSection"]/div[1]/div[3]/div[4]/div[1]/div[5]/div[1]')

    web.execute_script("arguments[0].scrollIntoView();", element) 

    # Take the screenshot -----------------------------
    time.sleep(4)
    
    screenshot(imgName)

    return ('done')

def final ():
    web.get('https://great-success.netlify.app')

    print("Script has ended. Please go and check your screenshots")
    
# --------Black Friesdays -------------------

def black_friesdays(page_url, imgName, x):

    web.get(page_url)
    
    time.sleep(5)

    work_on(x, imgName)
    # Accept cookies ------------------------------
    try:
        cookies_com()
    except:
        pass

    delay = 10
    try:
        myElem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.ID, 'left-ad-above-middle')))
    except TimeoutException:
        print ("Loading took too much time!")
        black_friesdays(page_url, imgName, x)

    time.sleep(4)

    element = web.find_element_by_xpath('//*[@id="left-ad-above-middle"]')

    # --------------------------------------------------------------
    # web.execute_script("arguments[0].scrollIntoView();", element) 
    # web.execute_script("window.scrollBy(0, -100)")
    # --------------------------------------------------------------

    # desired_y = (element.size['height'] / 2) + element.location['y']
    # current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    # scroll_y_by = desired_y - current_y
    # web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 
    # web.execute_script("window.scrollBy(0, -100)")

    # ----------------------------------------------------------------------
    actions = ActionChains(web)
    actions.move_to_element(element).perform()

    # web.execute_script("window.scrollBy(0, -100)")

    time.sleep(2)

    desired_y = (element.size['height'] / 2) + element.location['y']
    current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    scroll_y_by = desired_y - current_y + 100  # това +100 е под въпрос
    web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 


    # web.execute_script("window.scrollBy(0, -100)")

    # Take the screenshot -----------------------------
    time.sleep(4)
    
    screenshot(imgName)

    return ('done')

# -------------- Website Banner Premiun -------------------------------------

def website_banner_premium(imgName, x, src):


    web.get('https://de.ingrammicro.eu/')
    
    time.sleep(5)

    work_on(x, imgName)
    # Accept cookies ------------------------------
    try:
        cookies_eu()
    except:
        pass

    delay = 10

    num =1                                                                                  
    my_elem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[3]/div/span[1]/a/img')))
    checker = located_slot1(my_elem, imgName, src, num)

    if not checker:
       try:   
        num +=1                                                                                    
        my_elem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[3]/div/span[2]/a/img')))
        checker2 = located_slot1(my_elem, imgName, src, num)

        if not checker2:
            website_banner_premium(imgName, x, src)

       except TimeoutException:
        print ("Slot 2 is empty")     
 
    return ('done')

def located_slot1 (my_elem, imgName, src, num):

    source = my_elem.get_attribute("src")
    img = source.split("=")

    if src in img:

        premium_scroll_to()

        # Take the screenshot -----------------------------
        time.sleep(4)

        print("image found in slot " + str(num))

        screenshot(imgName) 
        
        bool = True

    else: 
        print("NOT in slot " + str(num))
        bool = False
    
    return bool

def premium_scroll_to():

    element = web.find_element_by_xpath('//*[@id="main"]/div[3]/div')

    # ----------------------------------------------------------------------
    actions = ActionChains(web)
    actions.move_to_element(element).perform()
    
    time.sleep(2)
    
    desired_y = (element.size['height'] / 2) + element.location['y']
    current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    scroll_y_by = desired_y - current_y + 100  # това +100 е под въпрос
    web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 


# -------------- Website Banner Standard -------------------------------------

def website_banner_standard(imgName, x, src):


    web.get('https://de.ingrammicro.eu/')
    
    time.sleep(5)

    work_on(x, imgName)
    # Accept cookies ------------------------------
    try:
        cookies_eu()
    except:
        pass

    delay = 10

    num =1                                                                                  
    my_elem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[4]/div/span[1]/a/img')))
    checker = located_slot(my_elem, imgName, src, num)

    if not checker:
       try:     
        num +=1                                                                             
        my_elem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[4]/div/span[2]/a/img')))
        checker2 = located_slot(my_elem, imgName, src, num)

        if not checker2:
            try:   
              num +=1                                                                                    
              my_elem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[4]/div/span[3]/a/img')))
              checker3 = located_slot(my_elem, imgName, src, num)

              if not checker3:
                  try:     
                    num +=1                                                                                  
                    my_elem = WebDriverWait(web, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[4]/div/span[4]/a/img')))
                    checker4 = located_slot(my_elem, imgName, src, num)

                    if not checker4:
                        website_banner_premium(imgName, x, src, num)

                  except TimeoutException:
                        print ("Slot 4 is empty")  

            except TimeoutException:
                 print ("Slot 3 is empty")  

       except TimeoutException:
            print ("Slot 2 is empty")     

    return ('done')

def located_slot (my_elem, imgName, src, num):

    source = my_elem.get_attribute("src")
    img = source.split("=")

    if src in img:

        standard_scroll_to()

        # Take the screenshot -----------------------------
        time.sleep(4)

        print("image found in slot " + str(num))

        screenshot(imgName) 
        
        bool = True

    else: 
        print("NOT in slot " + str(num))
        bool = False
    
    return bool

def standard_scroll_to():

    element = web.find_element_by_xpath('//*[@id="main"]/div[4]/div')

    # ----------------------------------------------------------------------
    actions = ActionChains(web)
    actions.move_to_element(element).perform()
    
    time.sleep(2)
    
    desired_y = (element.size['height'] / 2) + element.location['y']
    current_y = (web.execute_script('return window.innerHeight') / 2) + web.execute_script('return window.pageYOffset')
    scroll_y_by = desired_y - current_y + 100  # това +100 е под въпрос
    web.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by) 