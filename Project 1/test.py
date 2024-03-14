from tkinter import BOTH, LEFT, TOP, Button, Entry, Frame, Label, PhotoImage, StringVar, Tk,Radiobutton,StringVar,IntVar,filedialog
from idlelib.tooltip import Hovertip
from tkinter import messagebox
import sys
import os
import warnings
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
import time
import pandas as pd
from openpyxl import load_workbook
from urllib.parse import urlparse
import warnings
import numpy as np
from openpyxl.utils import get_column_letter,column_index_from_string
import requests
from zipfile import ZipFile
warnings.filterwarnings("ignore")


def test():    
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
    driver.get('https://cms.officeally.com/')
    print('done')    

def test2():
    options = Options()            
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument("--disable-popup-blocking")
    driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=options)
    driver.maximize_window()
    driver.get('https://cms.officeally.com')
    print('done') 

if __name__=="__main__":
    test2()