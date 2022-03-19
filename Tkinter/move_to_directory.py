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

import os 
import sys
import shutil
import zipfile

# Insert functions here -------------------------

# ---------------------------------------------------------
def move_to_dir():
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



   #  Homepage Banners
   for x in range(2, 160):

      # read value from cel A
      cel_A = 'A' + str(x) 
      name_value = sheet[cel_A]
      imgName = str(name_value.value) + week

      # read value from  cel D
      cel_D = 'D' + str(x) 
      page_value = sheet[cel_D]
      zip_path = str(page_value.value)


      if imgName != 'None':

         path = './Screenshots'
         name = imgName

         for root, dirs, files in os.walk(path):
            if name in files:

               path_to_file = './Screenshots/' + r"\\" + imgName

               zip = zipfile.ZipFile(zip_path,'a')
               zip.write(path_to_file, os.path.basename(path_to_file))
               print (imgName)
               zip.close()

