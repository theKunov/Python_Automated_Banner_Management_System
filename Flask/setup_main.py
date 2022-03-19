from selenium import webdriver
from webdriver_manager import driver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import string
# import pyautogui ??

from selenium.webdriver.common.action_chains import ActionChains
import unittest, win32com.client, time
from openpyxl import load_workbook
# import datetime ??

# Insert functions here -------------------------

from setup_functions import start, central_banner, middle_banner_search, middle_banner_details, skyscrapper
# ---------------------------------------------------------

class CB:
    Tower = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046418"
    Displays = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046420"
    Notebook = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046421"
    Fax = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046422"
    Server = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046423"
    Accessories = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046424"
    Cables = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046425"
    Network = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046426"
    Systems = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046427"
    Consumables = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046428"
    Electronics = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046429"
    Appliances = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046430"
    Optical = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046431"
    Devices = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046432"
    Lightning = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/264372642"
    LogisticServices = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/264372643"
    Processors = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046433"
    Mobility = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/264372644"
    Games = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046434"
    OfficeEquipment = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046445"
    Care = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046436"
    POS = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/139023758"
    Protection = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046437"
    Beamer = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046438"
    Camera = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046439"
    Hardware = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046440"
    Software = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046441"
    Storage = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046442"
    TelephonyEquipment= "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046443"
    WarrantiesServices= "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046444"

class MBS1:
    Tower = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046451"
    Displays = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046452"
    Notebook = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046453"
    Fax = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046454"
    Server = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046455"
    Accessories = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/189328566"
    Cables = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/264372674"
    Network = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046460"
    Systems = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046469"
    Consumables = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046492"
    Electronics = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/264372675"
    Appliances = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/264372676"
    Optical = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046461"
    Devices = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046493"
    Lightning = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/264372677"
    LogisticServices = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/264372678"
    Processors = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046494"
    Mobility = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/265007405"
    Games = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/209671618"
    OfficeEquipment = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046495"
    Care = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/265007406" 
    POS = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/265007407"
    Protection = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046462"
    Beamer = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046496"
    Camera = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046463"
    Hardware = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/178107276"
    Software = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046497"
    Storage = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046464"
    TelephonyEquipment= "http://de.ingrammicro.int/Site/ProductPlacement/Copy/265007408"
    WarrantiesServices= "http://de.ingrammicro.int/Site/ProductPlacement/Copy/265007409"  

class MBS2:
    Tower = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046456"
    Displays = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046457"
    Notebook = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046458"
    Fax = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046459"
    Server = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/160677654"
    Accessories = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/189328567"
    Cables = ""
    Network = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046472"
    Systems = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046478"
    Consumables = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/185580151"
    Electronics = ""
    Appliances = ""
    Optical = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046479"
    Devices = ""
    Lightning = ""
    LogisticServices = ""
    Processors = ""
    Mobility = ""
    Games = ""
    OfficeEquipment = ""
    Care = ""
    POS = ""
    Protection = ""
    Beamer = ""
    Camera = ""
    Hardware = ""
    Software = ""
    Storage = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046475"
    TelephonyEquipment= ""
    WarrantiesServices= ""  

class SKY:
    Apple = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046491"
    Cisco = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046480"
    Dell  = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046481"
    HP = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046489"
    Lenovo = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046482"
    Microsoft = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046490"
    Digital = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046483"
    Kingston = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046484"
    Supermicro = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046485"
    Seagate = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046486"
    Acer = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046488"
    Logitech = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/176046487"
    Netapp = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/188793909"
    Benq = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/249950110"
    Optoma = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/269445029"
    Netgear = "http://de.ingrammicro.int/Site/ProductPlacement/Copy/285241894"


failed_banners = []

# set file path
filepath="./DE Banner Setup.xlsx"

# load excel.xlsx
wb = load_workbook(filepath)

# activate demo.xlsx
sheet = wb.active

import sys, json
name = []

def run_prog():

    update()
#  Homepage Banners
    for x in range(2, 160):
        
       # read value from cel A
       cel_A = 'A' + str(x) 
       name_value = sheet[cel_A]
       imgName = str(name_value.value)
       # read value from  cel D (start date)
       cel_D = 'D' + str(x) 
       srt_date_value = sheet[cel_D]
       strDate = str(srt_date_value.value)
       # read value from  cel E (end date)
       cel_E = 'E' + str(x) 
       end_date_value = sheet[cel_E]
       endDate = str(end_date_value.value)
    
       # read value from  cel B (slot number)
       cel_B = 'B' + str(x) 
       slot_value = sheet[cel_B]
       slot = int(slot_value.value) #add "int" after test
    
       if x == 2 : 
        start()

        empty_cel_checker = str(name_value.value)

       if x != 1 and empty_cel_checker != 'None':
            name.append(imgName)
            sys.stdout = open('./static/set.js', 'w')
            jsonlist = json.dumps(name)
            print("var jsonName = '{}' ".format(jsonlist))
    
       if imgName != 'None':

            name_cat = imgName.split("-")

            # Central Banner (Premium & Standard) 
            if 'CB' in name_cat: 
              category = name_cat[-2] 
              if category == "Services":
               category = name_cat[-3] + name_cat[-2]
              elif category == "Equipment":
               category = name_cat[-3] + name_cat[-2]
              page_url = getattr(CB, category)
              central_banner(page_url, imgName, strDate, endDate, x)            
            # Spotlight Banner (Search page)
            elif 'MBS' in name_cat: 
               category = name_cat[-2] 
               if category == "Services":
                   category = name_cat[-3] + name_cat[-2]
               elif category == "Equipment":
                   category = name_cat[-3] + name_cat[-2]
               if 670 <= slot <= 674 or 680 <= slot <= 703 or slot == 826:    
                   try:
                       page_url = getattr(MBS1, category)
                       middle_banner_search(page_url, imgName, strDate, endDate, x)
                   except: 
                       # print('Category does NOT exist or something went wrong.\n Please set banner manually')
                       failed_banners.append(imgName)
               elif 675 <= slot <= 679 or 704 <= slot <= 727 or slot == 829:    
                   try:
                       page_url = getattr(MBS2, category)
                       middle_banner_search(page_url, imgName, strDate, endDate, x)
                   except: 
                       # print('Category does NOT exist or something went wrong.\n Please set banner manually')
                       failed_banners.append(imgName)           
            # Spotlight Banner (Product details)
            elif 'MBD' in name_cat:
               category = name_cat[-2] 
               if category == "Services":
                   category = name_cat[-3] + name_cat[-2]
               elif category == "Equipment":
                   category = name_cat[-3] + name_cat[-2]
               if 670 <= slot <= 674 or 680 <= slot <= 703 or slot == 826:    
                   try:
                       page_url = getattr(MBS1, category)
                       middle_banner_details(page_url, imgName, strDate, endDate, x)
                   except: 
                       # print('Category does NOT exist or something went wrong.\n Please set banner manually')
                       failed_banners.append(imgName)
               elif 675 <= slot <= 679 or 704 <= slot <= 727 or slot == 829:    
                   try:
                       page_url = getattr(MBS2, category)
                       middle_banner_details(page_url, imgName, strDate, endDate, x)
                   except: 
                       # print('Category does NOT exist or something went wrong.\n Please set banner manually')
                       failed_banners.append(imgName)           
            #Skyscrappers
            elif 'SKY' in name_cat:

               category = name_cat[2] 
               try:
                page_url = getattr(SKY, category)
                skyscrapper(page_url, imgName, strDate, endDate, x)
               except: 
               #  print('Category does NOT exist or something went wrong.\n Please set banner manually')
                failed_banners.append(imgName)

    if x == 159:
        # print('Failed banners (please set those manualy):') 

        for i in failed_banners:
            ""   
        #  print(i) # need to decide how to display the failed banners information!! 

   
# Update xlsx file -------------------------
def update():
    xlapp = win32com.client.DispatchEx("Excel.Application")

    # Must add full path to work
    file_path = r"C:\Users\bgkunb00\OneDrive - Ingram Micro\Desktop\Python\Flask - Test Update\DE Banner Setup.xlsx"

    wb = xlapp.Workbooks.Open(file_path)
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    xlapp.DisplayAlerts = False
    time.sleep(2)
    wb.Save()
    time.sleep(3)
    xlapp.Quit()
    wb.Close()
    time.sleep(1)
# ---------------------------------------------------------