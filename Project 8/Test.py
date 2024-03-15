from tkinter import BOTH, LEFT, TOP, Button, Entry, Frame, Label, PhotoImage, StringVar, Tk,Radiobutton,StringVar,IntVar,filedialog
from idlelib.tooltip import Hovertip
import sys
import os
from operator import itemgetter
import time
from tkinter import messagebox
from os import listdir
from os.path import isfile, join
from datetime import datetime, date
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, ElementClickInterceptedException, StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import openpyxl
import shutil
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from urllib.parse import urlparse
import warnings
import numpy as np
from openpyxl.utils import get_column_letter,column_index_from_string
from datetime import datetime
from pytz import timezone
import requests
from zipfile import ZipFile

warnings.filterwarnings("ignore")

global rows
global xpath
global heding
global status
global key
global nme
element_1 = None
global j

def process_1():        
    global element_1    
    global ans1
    global ans2 
    global driver
    global j
    global dob
    global mid
    global frm_dt
    global to_dt
    global cpt    
    
    try:
        options = Options()            
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_argument("--disable-popup-blocking")
        driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=options)
        driver.maximize_window()
        driver.get('https://secure.uhcprovider.com')
    except Exception as e:                        
        try:
            options = Options()            
            options.add_argument('--ignore-certificate-errors')
            options.add_argument('--ignore-ssl-errors')
            options.add_argument("--disable-popup-blocking")
            response = requests.get('https://chromedriver.storage.googleapis.com/LATEST_RELEASE')
            latest_version = response.text.strip()
            chrome_driver_url = f'https://chromedriver.storage.googleapis.com/{latest_version}/chromedriver_win32.zip'
            response = requests.get(chrome_driver_url)
            with open('chromedriver_win32.zip', 'wb') as f:
                f.write(response.content)
            with ZipFile('chromedriver_win32.zip', 'r') as zip_ref:
                zip_ref.extractall('.')
            driver_path = os.path.abspath('chromedriver.exe')
            driver = webdriver.Chrome(executable_path=driver_path,options=options)
            driver.maximize_window()
            driver.get('https://secure.uhcprovider.com')
        except Exception as e:
            messagebox.showinfo("Internet Problem","Pls Check Your Internet Connection")
            sys.exit(0)

    def click(xpath,heding,status):
        counter = 0
        while counter < 5:
            try:             
                WebDriverWait(driver, 0).until(EC.element_to_be_clickable((By.XPATH,xpath))).click()
                break
            except Exception as e:
                time.sleep(1)
                counter += 1
        else:
            print('1 - ' + status)
            raise e        
    try:
        xpath= "/html/body/div[2]/div/div[3]/div[1]/div[1]/header/div[2]/div/nav/div[1]/ul/li[2]/button"
        heding="Claims & Payments"
        status="Claims & Payments Button Not Found"
        click(xpath,heding,status)
        
        xpath= "/html/body/div[2]/div/div[3]/div[1]/div[1]/header/div[2]/div/nav/div[2]/div/div/ul/li[1]/div/ul/div[1]/li/a"
        heding="Look up a Claim"
        status="Look up a Claim Button Not Found"
        click(xpath,heding,status)  
    except Exception as e:
        print('Expection')
        xpath= "/html/body/div[1]/div/div[3]/div[1]/div[1]/header/div[2]/div/nav/div[1]/ul/li[2]/button"
        heding="Claims & Payments"
        status="Claims & Payments Button Not Found"
        click(xpath,heding,status)
        
        xpath= "/html/body/div[1]/div/div[3]/div[1]/div[1]/header/div[2]/div/nav/div[2]/div/div/ul/li[1]/div/ul/div[1]/li/a"
        heding="Look up a Claim"
        status="Look up a Claim Button Not Found"
        click(xpath,heding,status)  

if __name__=="__main__":        
    process_1()