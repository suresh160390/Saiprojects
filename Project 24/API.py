from tkinter import BOTH, LEFT, TOP, Button, Entry, Frame, Label, PhotoImage, StringVar, Tk,Radiobutton,StringVar,IntVar,filedialog
from idlelib.tooltip import Hovertip
from tkinter import ttk
from tkinter import messagebox
import pyodbc
import pandas as pd
import time
import pandas as pd
from sqlalchemy import create_engine
import subprocess
import platform
import logging
import os
import shutil
import sys

def combo():
   
    root=Tk()

    if getattr(sys, 'frozen', False):       
        image_path = os.path.join(sys._MEIPASS, 'Static', 'Close.png')
        image_path1 = os.path.join(sys._MEIPASS, 'Static', 'Mapping1.png')
    else:
        image_path = os.path.join(os.getcwd(), 'Static', 'Close.png')
        image_path1 = os.path.join(os.getcwd(), 'Static', 'Mapping1.png')

    root.title("SQL Server Connect & Data Extract")
    root.resizable(False,False)

    w = 500
    h = 170
    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))
    root.config(bg="#2c3e50",highlightbackground="blue",highlightthickness=1)    
    
    Frame1=Frame(root,bg="gold")
    Frame1.pack(side=TOP,fill=BOTH)
    title=Label(Frame1,text="SQL Server Data Extract Report",font=("Calibri",20,"bold","italic"),bg="gold",fg="black",justify="center")
    title.grid(row=0,columnspan=2,padx=8,pady=8)
    title.pack()
    
    Frame2=Frame(root,bg="#2c3e50")
    Frame2.place(x=0,y=40,width=500,height=300)
    
    title1=Label(Frame2,text="Select Required Report :",font=("Calibri",17,"bold","italic"),bg="#2c3e50",fg="white",justify=LEFT)
    title1.grid(row=0,column=0,padx=5,pady=5,sticky="W")
  
    answer=StringVar()
    answer.set("")

    title_3=Label(Frame2,text=answer.get(),textvariable=answer,font=("Calibri",12,"bold","italic"),bg="#2c3e50",fg="red",justify="center",width=65)
    title_3.grid(row=3,column=0,columnspan=2,padx=0,pady=0,sticky="W")
        
    options = ["New", "Dupilcate"]
    combo_box = ttk.Combobox(Frame2, values=options,textvariable="",width=29,justify="left",font=("Calibri",11,"bold","italic"))
    combo_box.grid(row=0,column=0,padx=265,pady=5,sticky="W")    
        
    def Click_Done():
        global ans

        ans=combo_box.get()

        if ans=="":
           answer.set("Dropdown Empty Is Not Allowed...")
        else:
            root.destroy()
            return ans
    
    photo1 = PhotoImage(file=image_path1)

    btn=Button(Frame2,command=Click_Done,text="Done",image=photo1,borderwidth=0,bg="#2c3e50")
    btn.grid(row=4,column=0,padx=115,pady=3,sticky="W")

    def Close():
        sys.exit(0)   
    
    photo = PhotoImage(file=image_path)   

    btn1=Button(Frame2,command=Close,text="Exit",image=photo,borderwidth=0,bg="#2c3e50")
    btn1.grid(row=4,column=0,padx=315,pady=3,sticky="W")

    def disable_event():
        pass

    myTip = Hovertip(btn,'Click to Done Continue Process',hover_delay=1000)
    myTip1 = Hovertip(btn1,'Click to Exit Process',hover_delay=1000)

    root.protocol("WM_DELETE_WINDOW", disable_event)

    root.mainloop()

def api_con():

    global ans
    combo()

    sec=ans 

    folder_create = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads/')
    # save_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads/Temp')
    save_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads\\Temp\\')
    # file = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads\\Temp\\Files\\')
        
    fld='Temp'
    
    if not os.path.isdir(save_path):
        os.mkdir(folder_create + 'Temp')        
    else:
        path1 = os.path.join(folder_create, fld)
        shutil.rmtree(path1)
        os.mkdir(folder_create + 'Temp')                

    server = 'thesnfistecmread2.database.windows.net'
    database = 'ecm_live'
    username = 'Ecmreaduser2'
    password = 'P4!@!5cF&mL'
    driver = 'ODBC Driver 17 for SQL Server'

    # connection_string = f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};UID={username};PWD={password};Trusted_Connection=yes;"
    connection_string = f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
    
    logging.basicConfig(level=logging.INFO) 

    try:
        connection = pyodbc.connect(connection_string)
        cursor = connection.cursor()

        sql_query = "SELECT encounterId, claimId, patientid FROM claims.claim WHERE [accountid] = 10 AND CONVERT(date, [firstbilleddate]) > '2023-10-01'"                        
        cursor.execute(sql_query)
        rows  = cursor.fetchall()
        df = pd.DataFrame(rows, columns=['encounterId', 'claimId', 'patientid'])        
        
        df.to_excel(save_path + 'output.xlsx', index=False)        
        cursor.close()
        connection.close()
        # df = pd.DataFrame([tuple(row) for row in rows], columns=[column[0] for column in cursor.description])
        # print(df)
        # time.sleep(120)
    except pyodbc.Error as e:
        print(f"Error connecting to the database: {e}")
        # time.sleep(60)   
    
    print('Completed')

# def test():    
#     server = 'thesnfistecmread2.database.windows.net'
#     database = 'ecm_live'
#     username = 'Ecmreaduser2'
#     password = 'P4!@!5cF&mL'
#     driver = 'ODBC Driver 17 for SQL Server'

#     # Define the connection string to your Power BI dataset
#     power_bi_connection_string = "DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}"

#     # Create SQLAlchemy engine
#     engine = create_engine(f"mssql+pyodbc:///?odbc_connect={power_bi_connection_string}")
    
#     df=pd.DataFrame()
#     # Export DataFrame to SQL Server
#     df.to_sql('YourPowerBITable', con=engine, if_exists='replace', index=False)

# def power_bi():
#     pbix_file_path = r'"C:\Users\sanandrao\Desktop\New folder (4)\bi.pbix"'
#     process = subprocess.Popen(['start', '""', pbix_file_path], shell=True)
    
#     time.sleep(10)

#     if platform.system() == 'Windows':
#         subprocess.run(['taskkill', '/F', '/T', '/PID', str(process.pid)], shell=True)
#     elif platform.system() == 'Linux':
#         subprocess.run(['kill', '-9', str(process.pid)], shell=True)
#     elif platform.system() == 'Darwin':  # macOS
#         subprocess.run(['kill', '-9', str(process.pid)], shell=True)
    

if __name__=="__main__":        
    combo()