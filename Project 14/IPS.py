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
from datetime import datetime, date
from pytz import timezone
import requests
from zipfile import ZipFile
import subprocess

warnings.filterwarnings("ignore")

global rows
global xpath
global heding
global status
global key
global nme
global element_1  
global j

def user_pass():
    global ans1
    global ans2
    root=Tk()

    if getattr(sys, 'frozen', False):       
        image_path = os.path.join(sys._MEIPASS, 'Static', 'Close.png')
        image_path1 = os.path.join(sys._MEIPASS, 'Static', 'Mapping1.png')
    else:
        image_path = os.path.join(os.getcwd(), 'Static', 'Close.png')
        image_path1 = os.path.join(os.getcwd(), 'Static', 'Mapping1.png')

    root.title("Sharepoint - User Login And File Details")
    root.resizable(False,False)

    w = 600
    h = 180
    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))
    root.config(bg="#2c3e50",highlightbackground="blue",highlightthickness=1)

    Frame1=Frame(root,bg="gold")
    Frame1.pack(side=TOP,fill=BOTH)
    title=Label(Frame1,text="User Name and Password Details...",font=("Calibri",20,"bold","italic"),bg="gold",fg="black",justify="center")
    title.grid(row=0,columnspan=2,padx=8,pady=8)
    title.pack() 
 
    Frame2=Frame(root,bg="#2c3e50")
    Frame2.place(x=0,y=40,width=698,height=150)    
    
    title3=Label(Frame2,text="User Name :",font=("Calibri",11,"bold","italic"),bg="#2c3e50",fg="white",justify=LEFT)
    title3.grid(row=0,column=0,padx=5,pady=5,sticky="W")
    
    txt1=Entry(Frame2,font=("Calibri",11,"bold","italic"),width=50,justify="left")
    txt1.grid(row=0,column=1,padx=30,pady=5,sticky="W")
    
    title4=Label(Frame2,text="Password :",font=("Calibri",11,"bold","italic"),bg="#2c3e50",fg="white",justify=LEFT)
    title4.grid(row=1,column=0,padx=5,pady=5,sticky="W")
    
    txt2=Entry(Frame2,font=("Calibri",11,"bold","italic"),width=50,justify="left")
    txt2.grid(row=1,column=1,padx=30,pady=5,sticky="W")

    answer=StringVar()
    answer.set("")

    def Click_Done():
        global ans1
        global ans2        
        
        ans1=txt1.get()
        ans2=txt2.get()        

        if ans1=="":
           answer.set("User Name Field Empty Is Not Allowed...")
        elif ans2=="":
            answer.set("Password Field Empty Is Not Allowed...")        
        else:    
            root.destroy()
            return ans1,ans2
                
    Frame4=Frame(root,bg="#2c3e50")
    Frame4.place(x=0,y=105,width=698,height=20)
    
    title_3=Label(Frame4,text=answer.get(),textvariable=answer,font=("Calibri",9,"bold","italic"),bg="#2c3e50",fg="Red",justify=LEFT)
    title_3.grid(row=0,column=0,columnspan=2,padx=200,pady=0,sticky="E")
    
    Frame3=Frame(root,bg="#2c3e50")
    Frame3.place(x=0,y=125,width=698,height=200)
       
    photo1 = PhotoImage(file=image_path1)
    
    btn1=Button(Frame3,command=Click_Done,text="Run",image=photo1,borderwidth=0,bg="#2c3e50")
    btn1.grid(row=2,column=0,padx=140,pady=0,sticky="W")

    def Close():
        sys.exit(0)

    photo = PhotoImage(file=image_path)    

    btn2=Button(Frame3,command=Close,text="Close",image=photo,borderwidth=0,bg="#2c3e50")
    btn2.grid(row=2,column=1,padx=100,pady=0,sticky="E")

    def disable_event():
        pass

    txt1.focus_set()

    myTip = Hovertip(btn1,'Click to Done Continue Process',hover_delay=1000)
    myTip1 = Hovertip(btn2,'Click to Exit Process',hover_delay=1000)

    root.protocol("WM_DELETE_WINDOW", disable_event)
            
    root.mainloop()

def process():        
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

    user_pass()
                 
    user = ans1
    password = ans2

    fld_path1 = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads/')
    check_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads/Temp')
    down_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads\\Temp\\')
    # file = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads\\Temp\\Files\\')
    
    # fld_path = os.path.abspath(os.path.dirname(__file__))
    # fld_path1=fld_path + '/'
    # check_path = fld_path + '/Temp'
    # down_path = fld_path + '/Temp/'
    # down_path = down_path.replace('/', '\\')
    
    fld='Temp'
    
    if not os.path.isdir(check_path):
        os.mkdir(fld_path1 + 'Temp')        
    else:
        path1 = os.path.join(fld_path1, fld)
        shutil.rmtree(path1)
        os.mkdir(fld_path1 + 'Temp')        

    try:
        options = Options()            
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        prefs = {"download.default_directory" : down_path,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True}
        options.add_experimental_option("prefs",prefs)
        driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=options)
        driver.maximize_window()
        driver.get('https://www.therapynotes.com/app/login/Integrated30909/')            
    except Exception as e:
        try:
            options =  webdriver.ChromeOptions()        
            options.add_argument('--ignore-certificate-errors')
            options.add_argument('--ignore-ssl-errors')
            options.add_argument("--disable-popup-blocking")
            prefs = {"download.default_directory" : down_path,
                    "download.prompt_for_download": False,
                    "download.directory_upgrade": True,
                    "safebrowsing.enabled": False,
                    "download.extensions_to_open": "",
                    "profile.default_content_settings.popups": 0,
                    "profile.content_settings.exceptions.automatic_downloads.www.therapynotes.com.setting": 1}
            
            options.add_experimental_option("prefs",prefs)                
            driver_path = os.path.abspath('chromedriver.exe')                
            driver = webdriver.Chrome(executable_path=driver_path,options=options)                
            driver.maximize_window()
            driver.get('https://www.therapynotes.com/app/login/Integrated30909/')                                         
        except Exception as e:
            messagebox.showinfo("Driver Problem","Pls Check Your Driver Version")
            sys.exit(0)
        
    def text_box(xpath,heding,status,key):                
        counter = 0
        while counter < 30:
            try:   
                WebDriverWait(driver, 0).until(EC.visibility_of_element_located((By.XPATH,xpath))).clear()
                WebDriverWait(driver, 0).until(EC.visibility_of_element_located((By.XPATH,xpath))).send_keys(key)
                break
            except Exception as e:
                time.sleep(1)
                counter += 1
        else:
            # print('Moved Exception')
            # raise e
            messagebox.showinfo(heding, status)
            sys.exit(0)        
    
    def text_box_key(xpath,heding,status,key):                
        counter = 0
        while counter < 30:
            try:                   
                input_field =WebDriverWait(driver, 0).until(EC.visibility_of_element_located((By.XPATH,xpath)))
                input_field.send_keys(Keys.CONTROL + "a") 
                input_field.send_keys(Keys.BACKSPACE)                 
                input_field.send_keys(key)               
                input_field.send_keys(Keys.TAB)
                break
            except Exception as e:
                time.sleep(1)
                counter += 1
        else:
            # print('Moved Exception')
            # raise e
            messagebox.showinfo(heding, status)
            sys.exit(0)    

    def text_box_js(xpath,heding,status,key):                
        counter = 0
        while counter < 30:
            try:
                js_code = "document.querySelector('" + xpath + "').setAttribute('value', '" + key + "');"
                WebDriverWait(driver, 0).until(lambda driver: driver.execute_script(js_code))               
                break
            except Exception as e:
                time.sleep(1)
                counter += 1
        else:
            messagebox.showinfo(heding, status)
            sys.exit(0)       

    def click(xpath,heding,status):
        counter = 0
        while counter < 30:
            try:             
                WebDriverWait(driver, 0).until(EC.element_to_be_clickable((By.XPATH,xpath))).click()
                break
            except Exception as e:
                time.sleep(1)
                counter += 1
        else:
            # print('Moved Exception')
            # raise e
            messagebox.showinfo(heding,status)
            sys.exit(0)                                

    def count(xpath,heding,status):
        global rows
        counter = 0
        while counter < 30:
            try:             
                rows=len(WebDriverWait(driver, 0).until(EC.presence_of_all_elements_located((By.XPATH,xpath))))
                break
            except Exception as e:
                time.sleep(1)
                counter += 1
        else:
            # print('Moved Exception')
            # raise e
            messagebox.showinfo(heding,status)
            sys.exit(0)                         

    def Alert():
        counter = 0
        while counter < 30:
            try:             
                WebDriverWait(driver, 0).until (EC.alert_is_present())
                a=driver.switch_to.alert
                a.accept()
                break
            except Exception as e:
                time.sleep(1)
                counter += 1
        else:
            messagebox.showinfo('Alert','Alert Not Present')                                                                

    xpath= '/html/body/div[4]/main/div/div/div[2]/div[4]/div[2]/div[1]/div/form/div[3]/div/input'
    heding="User Name"
    status="User Name Field Not Found"
    key=user
    text_box(xpath,heding,status,key)
    
    xpath= '/html/body/div[4]/main/div/div/div[2]/div[4]/div[2]/div[1]/div/form/div[4]/input'
    heding="Password"
    status="Password Field Not Found"
    key=password
    text_box(xpath,heding,status,key)

    xpath= "/html/body/div[4]/main/div/div/div[2]/div[4]/div[2]/div[1]/div/form/div[5]/button"
    heding="Log In"
    status="Log In Button Not Found"
    click(xpath,heding,status)

    # while True:
    #         xpath='/html/body/div[4]/main/div/div/div[2]/div[4]/div[2]/div[1]/div/form/div[1]/div/span[2]'         
    #         try:
    #             element_3 = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.XPATH, xpath))).text
    #             if element_3.lstrip().rstrip()=='The username and password you entered did not match any account.':                               
    #                 messagebox.showinfo('Login Status','Please Check - UserName or Password Wrong')                    
    #                 user_pass()
                           
    #                 xpath= '/html/body/div[4]/main/div/div/div[2]/div[4]/div[2]/div[1]/div/form/div[3]/div/input'
    #                 heding="User Name"
    #                 status="User Name Field Not Found"
    #                 key=ans1
    #                 text_box_key(xpath,heding,status,key)
                    
    #                 xpath= '/html/body/div[4]/main/div/div/div[2]/div[4]/div[2]/div[1]/div/form/div[4]/input'
    #                 heding="Password"
    #                 status="Password Field Not Found"
    #                 key=ans2
    #                 text_box(xpath,heding,status,key)

    #                 xpath= "/html/body/div[4]/main/div/div/div[2]/div[4]/div[2]/div[1]/div/form/div[5]/button"
    #                 heding="Log In"
    #                 status="Log In Button Not Found"
    #                 click(xpath,heding,status)
    #             else:
    #                 break    
    #         except Exception as e:                
    #                 break    
    
    xpath= "/html/body/form/div[4]/div/div/div/div[3]/ul/li"
    heding="Billing Table Count"
    status="Billing Table Count Not Found"        
    count(xpath,heding,status) 
    
    j=1
    while j<rows+1: 
        xpath= "/html/body/form/div[4]/div/div/div/div[3]/ul/li[{}]"                                
        cnm=WebDriverWait(driver, 0).until(EC.visibility_of_element_located((By.XPATH,xpath.format(j)))).text     
        if cnm.lstrip().rstrip()=='Billing':
            xpath= "/html/body/form/div[4]/div/div/div/div[3]/ul/li[{}]/a"   
            WebDriverWait(driver, 0).until(EC.element_to_be_clickable((By.XPATH,xpath.format(j)))).click()              
            break
        j=j+1
    
    xpath= "/html/body/form/div[5]/div/main/div[2]/div/div[3]/div/div[1]/div[2]/div[1]/div/div[1]/div[2]/div[2]/select/optgroup[1]/option"
    heding="Type Count"
    status="Type Count Not Found"        
    count(xpath,heding,status)

    j=1
    while j<rows+1: 
        xpath= "/html/body/form/div[5]/div/main/div[2]/div/div[3]/div/div[1]/div[2]/div[1]/div/div[1]/div[2]/div[2]/select/optgroup[1]/option[{}]"                                
        cnm=WebDriverWait(driver, 0).until(EC.visibility_of_element_located((By.XPATH,xpath.format(j)))).text     
        if cnm.lstrip().rstrip()=='Appointment with Completed Note':                
            WebDriverWait(driver, 0).until(EC.element_to_be_clickable((By.XPATH,xpath.format(j)))).click()              
            break
        j=j+1

    xpath= "/html/body/form/div[5]/div/main/div[2]/div/div[3]/div/div[1]/div[2]/div[1]/div/div[1]/div[4]/div[2]/span/select/option"
    heding="Date Count"
    status="Date Count Not Found"        
    count(xpath,heding,status)
    
    j=1
    while j<rows+1: 
        xpath= "/html/body/form/div[5]/div/main/div[2]/div/div[3]/div/div[1]/div[2]/div[1]/div/div[1]/div[4]/div[2]/span/select/option[{}]"                               
        cnm=WebDriverWait(driver, 0).until(EC.visibility_of_element_located((By.XPATH,xpath.format(j)))).text     
        if cnm.lstrip().rstrip()=='Last 30 Days':                
            WebDriverWait(driver, 0).until(EC.element_to_be_clickable((By.XPATH,xpath.format(j)))).click()              
            break
        j=j+1    
            
    xpath= "/html/body/form/div[5]/div/main/div[2]/div/div[3]/div/div[1]/div[2]/div[1]/div/div[1]/div[4]/div[2]/span/span[2]/span[1]/input"
    heding="From Date"
    status="From Date Field Not Found"
    key='7/17/2023'
    text_box_key(xpath,heding,status,key)  
    
    to_dt = date.today()    
    dob3 = to_dt.strftime("%m/%d/%Y") 

    xpath= "/html/body/form/div[5]/div/main/div[2]/div/div[3]/div/div[1]/div[2]/div[1]/div/div[1]/div[4]/div[2]/span/span[2]/span[2]/input"
    heding="To Date"
    status="To Date Field Not Found"
    key=dob3
    text_box_key(xpath,heding,status,key)

    xpath= "/html/body/form/div[5]/div/main/div[2]/div/div[3]/div/div[1]/div[2]/div[2]/span/input"
    heding="Search Click"
    status="Search Click Button Not Found"
    click(xpath,heding,status)   

    counter = 0
    while counter < 30:
        try:                   
            st = WebDriverWait(driver, 0).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div[5]/div/main/div[2]/div/div[3]/div/div[2]/div/div/div[1]/a'))).text                                  
            if st.lstrip().rstrip()=='Export Spreadsheet':
                WebDriverWait(driver, 0).until(EC.element_to_be_clickable((By.XPATH,'/html/body/form/div[5]/div/main/div[2]/div/div[3]/div/div[2]/div/div/div[1]/a'))).click() 
                # driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                # driver.switch_to.window(driver.window_handles[-1])
                break
        except Exception as e:                         
            time.sleep(1)
            counter += 1
    else:
        messagebox.showinfo('Export Table Loding', 'Loding Time Taken Very Long Time - Pls Try Some Time')
        sys.exit(0)

    counter = 0
    while counter < 30:
        try:                   
            st = WebDriverWait(driver, 0).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div[3]/div/div/div/div/a/div[2]'))).text                                  
            if st.lstrip().rstrip()=='Ready: Click to Download':
                WebDriverWait(driver, 0).until(EC.element_to_be_clickable((By.XPATH,'/html/body/form/div[3]/div/div/div/div/a/div[2]'))).click() 
                break
        except Exception as e:                         
            time.sleep(1)
            counter += 1
    else:
        messagebox.showinfo('Download Error', "Not Ready: Click to Download File - Pls Try Some Time")
        sys.exit(0)  

    counter=1
    while counter < 5:                    
        file1 = [f for f in listdir(down_path) if isfile(join(down_path, f))]
        lol_string1 = ' '.join(file1)
        lol_string2=lol_string1.split('.')[-1]                                                                                
        if lol_string2!='' and lol_string2!='tmp' and lol_string2 !="crdownload" and lol_string2 !="temp":
            time.sleep(1)
            file2 = [f for f in os.listdir(down_path) if os.path.isfile(os.path.join(down_path, f))]
            if file2:                                                     
                break
        else:
            time.sleep(1)    
            counter += 1
    else:
        messagebox.showinfo("Download Status", "File Not Download")
        sys.exit(0)             
    
    file2 = [f for f in os.listdir(down_path) if os.path.isfile(os.path.join(down_path, f))]
    lol_string3 = os.path.join(down_path, file2[0])  
    
    file=pd.read_excel(lol_string3,header=0)
    
    lst=['Date','Service Code','POS','Units','Last Name','First Name','DOB','Patient Member ID','Clinician Name','Location','Primary Insurer Name','Primary Diagnosis','Patient Balance Status','Modifier Code 1','Modifier Code 2','Modifier Code 3','Modifier Code 4']

    fin=file[['Date','Service Code','POS','Units','Last Name','First Name','DOB','Patient Member ID','Clinician Name','Location','Primary Insurer Name','Primary Diagnosis','Patient Balance Status','Modifier Code 1','Modifier Code 2','Modifier Code 3','Modifier Code 4']]    

    fin['Date'] = pd.to_datetime(fin['Date'].dt.date)
    fin['Date'] = pd.to_datetime(fin['Date']).dt.strftime('%m/%d/%Y')
    
    fin=pd.DataFrame(fin)

    for col in lst:
        if fin[col].dtype == 'object':
            fin[col] = fin[col].apply(lambda x: ''.join(e for e in str(x) if e.isalnum() or e.isspace())).str.strip()
            # fin[col] = fin[col].apply(lambda x: x.str.replace(r'[^a-zA-Z0-9]+', '', regex=True).str.strip())
    
    # fin = fin.apply(lambda x: x.str.replace(r'[^a-zA-Z0-9]+', '', regex=True).str.strip())

    # fin = fin[(fin['Date'].notna() & fin['Date'] != '') & (fin['Service Code'].notna() & fin['Service Code'] != '') & (fin['POS'].notna() & fin['POS'] != '') & (fin['Units'].notna() & fin['Units'] != '') & (fin['Last Name'].notna() & fin['Last Name'] != '') & (fin['First Name'].notna() & fin['First Name'] != '') & (fin['DOB'].notna() & fin['DOB'] != '') & (fin['Patient Member ID'].notna() & fin['Patient Member ID'] != '') & (fin['Clinician Name'].notna() & fin['Clinician Name'] != '') & (fin['Location'].notna() & fin['Location'] != '') & (fin['Primary Insurer Name'].notna() & fin['Primary Insurer Name'] != '') & (fin['Primary Diagnosis'].notna() & fin['Primary Diagnosis'] != '') & (fin['Patient Balance Status'].notna() & fin['Patient Balance Status'] != '') & (fin['Modifier Code 1'].notna() & fin['Modifier Code 1'] != '') & (fin['Modifier Code 2'].notna() & fin['Modifier Code 2'] != '') & (fin['Modifier Code 3'].notna() & fin['Modifier Code 3'] != '') & (fin['Modifier Code 4'].notna() & fin['Modifier Code 4'] != '')]
    
    fin = fin[(fin['Date'].notna()) & (fin['Service Code'].notna()) & (fin['POS'].notna()) & (fin['Units'].notna()) & (fin['Last Name'].notna()) & (fin['First Name'].notna()) & (fin['DOB'].notna()) & (fin['Patient Member ID'].notna()) & (fin['Clinician Name'].notna()) & (fin['Location'].notna()) & (fin['Primary Insurer Name'].notna()) & (fin['Primary Diagnosis'].notna()) & (fin['Patient Balance Status'].notna()) & (fin['Modifier Code 1'].notna()) & (fin['Modifier Code 2'].notna()) & (fin['Modifier Code 3'].notna()) & (fin['Modifier Code 4'].notna())]

    fin = fin[(fin['Date']!='nan') & (fin['Service Code']!='nan') & (fin['POS']!='nan') & (fin['Units']!='nan') & (fin['Last Name']!='nan') & (fin['First Name']!='nan') & (fin['DOB']!='nan') & (fin['Patient Member ID']!='nan') & (fin['Clinician Name']!='nan') & (fin['Location']!='nan') & (fin['Primary Insurer Name']!='nan') & (fin['Primary Diagnosis']!='nan') & (fin['Patient Balance Status']!='nan') & (fin['Modifier Code 1']!='nan') & (fin['Modifier Code 2']!='nan') & (fin['Modifier Code 3']!='nan') & (fin['Modifier Code 4']!='nan')]

    # fin = fin.dropna(subset=lst, how='all')

    # fin = fin.dropna()

    # fin = fin.applymap(lambda x: np.nan if x == 'NaN' else x)

    # fin = fin.replace('NaN', np.nan)
    # fin.to_excel(lol_string3 + 'Output' + '.xlsx', sheet_name='Data',index=False)
    print(fin)

    # filtered_df = fin[fin[lst].apply(lambda x: ~x.isna()).any(axis=1)]

    # print(filtered_df)

    # directory, file_without_extension = os.path.split(lol_string3)
    # file_name = os.path.splitext(file_without_extension)[0]

    # xlsx_file_path = os.path.join(file, f'{file_name}.xlsx')

    driver.quit()
    messagebox.showinfo("Process Status", "Process Completed")
    sys.exit(0)    

if __name__=="__main__":        
    process()