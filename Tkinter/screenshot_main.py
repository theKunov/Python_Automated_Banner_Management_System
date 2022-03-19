from selenium import webdriver
from webdriver_manager import driver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import string
import pyautogui

from selenium.webdriver.common.action_chains import ActionChains
import unittest
from openpyxl import load_workbook
import datetime

# Insert functions here -------------------------
from screenshots_functions import sponsored_news, sponsored_events, spotlight_home_banner, top_home, xl_slider_banner_1, xl_slider_banner_2, xl_slider_banner_3, uberall_banner_1, uberall_banner_2, uberall_banner_3, top_banner_paket, top_banner_paket_1
from screenshots_functions import central_banner, spotlight_banner_premium_1, spotlight_banner_details, skyscrapper_A, final, black_friesdays, website_banner_premium, website_banner_standard
# ---------------------------------------------------------

# get current week of the year + file extention
my_date = datetime.date.today()
year, week_number, day_of_week = my_date.isocalendar()

if week_number < 10:
   week = '-CW0' + str(week_number) + '.png'
else:
   week = '-CW' + str(week_number) + '.png'
    
# set file path
filepath="./DE Banner Screenshots.xlsx"

# load excel.xlsx
wb = load_workbook(filepath)

# activate demo.xlsx
sheet = wb.active

def screenshot():

   #Homepage Banners
   for x in range(24, 160):

      # read value from cel A
      cel_A = 'A' + str(x) 
      name_value = sheet[cel_A]
      imgName = str(name_value.value) + week
      # read value from  cel B
      cel_B = 'B' + str(x) 
      page_value = sheet[cel_B]
      page_url = str(page_value.value)

      # read value from  cel C
      cel_C = 'C' + str(x) 
      source_value = sheet[cel_C]
      src = str(source_value.value) + "&sys"

      if imgName != 'None':

         name_cat = imgName.split("-")

         # website banner premium
         if 'Premium' in name_cat and 'HME' in name_cat:
            website_banner_premium(imgName, x, src)

         # website banner premium
         elif 'Standard' in name_cat and 'HME' in name_cat:
            website_banner_standard(imgName, x, src)

         # top home banner
         elif 'Top' in name_cat and 'TSH' in name_cat:
            top_home(page_url, imgName, x)

         # spotlight home banner
         elif 'Spotlight' in name_cat:
            spotlight_home_banner(page_url, imgName, x)

         #  XL Slider banner (1, 2, 3)
         elif 'XL' in name_cat and 'Slider' in name_cat:

             if (imgName.find('1-IMO') != -1):
                 xl_slider_banner_1(page_url, imgName, x)

             if (imgName.find('2-IMO') != -1):
                 xl_slider_banner_2(page_url, imgName, x)

             if (imgName.find('3-IMO') != -1):
                 xl_slider_banner_3(page_url, imgName, x)

         # Uberall BBH (homepage)
         elif 'BBH' in name_cat:
             uberall_banner_3(imgName, x)

         # Uberall BBD 
         elif 'BBD' in name_cat:
            uberall_banner_2(imgName, x)

         # sponsored news
         elif 'Sponsored' in name_cat and 'News' in name_cat:
            sponsored_news(page_url, imgName, x)

         # sponsored event
         elif 'Sponsored' in name_cat and 'Event' in name_cat:
            sponsored_events(page_url, imgName, x)

         # Uberall BBS
         elif 'Uberall' in name_cat and 'BBS' in name_cat:
            uberall_banner_1(imgName, x)

         # Top Banner IMOS
         elif 'Top' in name_cat and 'Banner' in name_cat and 'IMOS' in name_cat:
            top_banner_paket(page_url, imgName, x)

         # Top Banner Paket IMOD
         elif 'Top' in name_cat and 'Banner' in name_cat and 'IMOD' in name_cat:
            top_banner_paket_1(page_url, imgName, x)

         # Central Banner (Premium & Standard) 
         elif 'CB' in name_cat:   
             central_banner(page_url, imgName, x)

         # Spotlight Banner Premium & Standard (1 & 2)
         elif 'MBS' in name_cat: 
             spotlight_banner_premium_1(page_url, imgName, x)

         # Spotlight Banner Premium and Standard (1 & 2)
         elif 'MBD' in name_cat:
            spotlight_banner_details(page_url, imgName, x)

         #Skyscrappers
         elif 'SKY' in name_cat:
            skyscrapper_A(page_url, imgName, x)


         # Black Friesdays
         elif 'Friesday' in name_cat:
            black_friesdays(page_url, imgName, x)

      #Final 
      if x == 159 :
         final()    
