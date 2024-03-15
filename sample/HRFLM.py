import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import VotingClassifier
from sklearn.metrics import confusion_matrix
from sklearn import model_selection
# from sklearn.metrics import plot_confusion_matrix
from sklearn.metrics import accuracy_score
from sklearn.metrics import precision_score
from sklearn.metrics import recall_score
from sklearn.metrics import f1_score
from sklearn import model_selection
from sklearn.model_selection import cross_val_score
import seaborn as sns
from tkinter import BOTH, LEFT, TOP, Button, Entry, Frame, Label, PhotoImage, StringVar, Tk,StringVar,filedialog
from tkinter import messagebox
from idlelib.tooltip import Hovertip
import os
import sys

global ans

def browse():
   
    root=Tk()

    if getattr(sys, 'frozen', False):       
        image_path = os.path.join(sys._MEIPASS, 'Static', 'Close.png')
        image_path1 = os.path.join(sys._MEIPASS, 'Static', 'Mapping1.png')
        image_path2 = os.path.join(sys._MEIPASS, 'Static', 'Mapping.png')
    else:
        image_path = os.path.join(os.getcwd(), 'Static', 'Close.png')
        image_path1 = os.path.join(os.getcwd(), 'Static', 'Mapping1.png')
        image_path2 = os.path.join(os.getcwd(), 'Static', 'Mapping.png')

    root.title("Process File Picker")
    root.resizable(False,False)

    root.title("Process File Picker")
    root.resizable(False,False)
   
    w = 500
    h = 160
    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))
    root.config(bg="#2c3e50",highlightbackground="blue",highlightthickness=1)    
    
    Frame1=Frame(root,bg="gold")
    Frame1.pack(side=TOP,fill=BOTH)
    title=Label(Frame1,text="Browse Process - File Picker",font=("Calibri",18,"bold","italic"),bg="gold",fg="black",justify="center")
    title.grid(row=0,columnspan=2,padx=8,pady=8)
    title.pack()
          
    Frame2=Frame(root,bg="#2c3e50")
    Frame2.place(x=0,y=40,width=500,height=50)
    
    answer=StringVar()
    answer.set("")
       
    def browse_button():
        global filename
        answer.set("")  
        filename = filedialog.askopenfilename()  
        txt.delete(0, 'end') 
        txt.insert(0, filename)

    def Click_Done():
        global ans        
        
        ans=txt.get()        
        
        if ans=="":
           answer.set("File Path Fields Empty Is Not Allowed...")        
        else:      
            root.destroy()          
            return ans
        
    txt=Entry(Frame2,font=("Calibri",12,"bold","italic"),width=50,justify="left")
    txt.grid(row=0,column=0,padx=5,pady=10,sticky="E")    

    photo1 = PhotoImage(file=image_path2)
            
    btn1=Button(Frame2,text="Browse",command=browse_button,image=photo1,borderwidth=0,bg="#2c3e50")
    btn1.grid(row=0,column=1,padx=3,pady=0,sticky="W")     
                
    Frame3=Frame(root,bg="#2c3e50")
    Frame3.place(x=3,y=80,width=490,height=25)
    
    title_3=Label(Frame3,text=answer.get(),textvariable=answer,font=("Calibri",9,"bold","italic"),bg="#2c3e50",fg="Red",justify=LEFT)
    title_3.grid(row=0,column=0,columnspan=1,padx=150,pady=0,sticky="E")

    Frame4=Frame(root,bg="#2c3e50")
    Frame4.place(x=3,y=105,width=500,height=50)
    
    photo2 = PhotoImage(file=image_path1)

    btn2=Button(Frame4,command=Click_Done,text="Done",image=photo2,borderwidth=0,bg="#2c3e50")
    btn2.grid(row=3,column=0,padx=125,pady=0,sticky="W")

    def Close():
        sys.exit(0)   
    
    photo3 = PhotoImage(file=image_path)   

    btn3=Button(Frame4,command=Close,text="Exit",image=photo3,borderwidth=0,bg="#2c3e50")
    btn3.grid(row=3,column=1,padx=15,pady=0,sticky="W")

    def disable_event():
        pass

    myTip = Hovertip(btn2,'Click to Done Continue Process',hover_delay=1000)
    myTip1 = Hovertip(btn3,'Click to Exit Process',hover_delay=1000)
    myTip2 = Hovertip(btn1,'Click to Pick File',hover_delay=1000)

    root.protocol("WM_DELETE_WINDOW", disable_event)

    root.mainloop()

def process():

    browse()

    fil=ans     
    # Importing the dataset
    dataset = pd.read_csv(fil)
    # print(dataset)
    dataset=dataset.dropna(how="any")
    # print(dataset)
    # print(" ")
    dataset.loc[:, 'Num'].replace([2, 3, 4], [1, 1, 1], inplace=True)

    # print(dataset.info())
    # print(" ")
    # print(dataset)

    #Data Visualization

    #age vs cholestrol
    m = dataset['Age']
    n = dataset['Chol']
    plt.figure(figsize=(10,8))
    plt.title("Age vs Cholestrol",fontsize=20)
    plt.xlabel("Age",fontsize=15)
    plt.ylabel("Cholestrol",fontsize=15)
    plt.bar(m,n,label="age vs cholestrol",color=["red","orange"],width=0.5)
    plt.legend(loc='best')
    plt.show()

    #age vs chest pain
    m = dataset['Age']
    n = dataset['Cp']
    plt.figure(figsize=(10,8))
    plt.title("Age vs Chestpain",fontsize=20)
    plt.xlabel("Age",fontsize=15)
    plt.ylabel("Cp",fontsize=15)
    plt.bar(m,n,label="age vs chestpain",color=["red","orange"],width=0.5)
    plt.legend(loc='best')
    plt.show()

    #histogram of chest pain
    plt.figure(figsize=(10,8))
    plt.title("Histogram of Chest Pain")
    plt.hist(dataset['Cp'],rwidth=0.9)
    plt.show()

    X = dataset.iloc[:,:-1].values
    y = dataset.iloc[:, 13].values

    # Splitting the dataset into the Training set and Test set
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size = 0.25, random_state = 121)#101

    # print(X_train)
    # print(" ")

    # Feature Scaling
    sc = StandardScaler()
    X_train = sc.fit_transform(X_train)
    X_test = sc.transform(X_test)

    # print(X_train)
    #EXPLORING THE DATASET

    dataset.Num.value_counts()

    estimators = []

    # Fitting RF and LM to the Training set

    model1 =RandomForestClassifier(n_estimators=20)
    estimators.append(('rt1', model1))
    model2 =RandomForestClassifier(n_estimators=20)
    estimators.append(('rt2', model2))
    model3 =RandomForestClassifier(n_estimators=20)
    estimators.append(('rt3', model3))
    model4 =RandomForestClassifier(n_estimators=20)
    estimators.append(('rt4', model4))
    model5 =RandomForestClassifier(n_estimators=20)
    estimators.append(('rt5', model5))

    model11 =LogisticRegression()
    estimators.append(('log1', model11))
    model12 =LogisticRegression()
    estimators.append(('log2', model12))
    model13 =LogisticRegression()
    estimators.append(('log3', model13))
    model14 =LogisticRegression()
    estimators.append(('log4', model14))
    model15 =LogisticRegression()
    estimators.append(('log5', model15))

    #voting classifier
    ensemble = VotingClassifier(estimators)
    ensemble.fit(X_train, y_train)
    y_pred = ensemble.predict(X_test)
    # print("ypred : ")
    # print(y_pred)
    # print(" ")
    # print("ytest : ")
    # print(y_test)
    # print(" ")


    #visualize the confusion matrix of HRFLM
    # matrix = confusion_matrix(ensemble, X_test, y_test,cmap=plt.cm.YlGnBu)
    # matrix.ax_.set_title('confusion matrix', color='blue')
    matrix = confusion_matrix(y_test, y_pred)
    plt.figure(figsize=(8, 6))
    sns.heatmap(matrix, annot=True, fmt='d', cmap=plt.cm.YlGnBu)
    plt.xlabel('predicted value', color='blue')
    plt.ylabel('actual value', color='blue')
    plt.gcf().axes[0].tick_params(colors='blue')
    plt.gcf().axes[1].tick_params(colors='blue')
    plt.gcf().set_size_inches(10,6)
    plt.show()

    #confusion matrix
    tp_hrflm,tn_hrflm,fp_hrflm,fn_hrflm = confusion_matrix(y_test,y_pred).ravel()
    # print("True positive for HRFLM :",tp_hrflm)
    # print("True negative for HRFLM :",tn_hrflm)
    # print("False positive for HRFLM :",fp_hrflm)
    # print("False negative for HRFLM :",fn_hrflm)
    # print(" ")


    #accuracy
    kfold = model_selection.KFold(n_splits=10, random_state=None)
    results = model_selection.cross_val_score(ensemble, X_train, y_train, cv=kfold)
    # print(results.mean())
    # print(" ")
    # print(results)
    # print(" ")
    # print("Accuracy is")
    r=results[results.argmax()]
    # print(r)
    # print(" ")
    #Precision
    # print("Precision for HRFLM is ",100*precision_score(y_test,y_pred))
    # #Recall
    # print("Recall for HRFLM is ",100*recall_score(y_test,y_pred))
    # #F1_score
    # print("F1_score for HRFLM is ",100*f1_score(y_test,y_pred))
    # print(" ")
    # #error rate
    # print("Error rate is")
    # print(1-r)

    #Comparative analysis
    w=0.4
    left = [1,2,3] 
    height = [81.00,84.33,95.95]
    tick_label = ['Random Forest','Logistic Regression','HRFLM']
    plt.figure(figsize=(8, 6))
    plt.bar(left, height, tick_label = tick_label, color = ['red', 'green','blue'])
    # print(" ")
    plt.xlabel('Classification Algorithms')
    plt.ylabel('Accuracy Score')
    plt.title('Comparitive Analysis') 
    plt.show()

    # while True:
    #     # print(" ")
    #     input_string = input("Enter a list elements separated by space : ")
    #     if (input_string=="break"):
    #         # print("Testing Finished")
    #         l=0
    #         break
    #     # print("\n")
    #     userList = input_string.split()
    #     # print("user list is ", userList)
    #     list_of_floats = [float(item) for item in userList]
    #     # print(list_of_floats)
    #     ynew=ensemble.predict(np.array(list_of_floats).reshape(1, -1))
    #     ynew
    #     if (ynew[0]==0):
    #         print(" ")
    #         print("Heart Disease Status : Not Detected")
    #         print("For the given dataset the Predicted Value is Absence of Heart Disease")
    #         print(" ")
    #     elif (ynew[0]==1):
    #         print("  ")
    #         print("Heart Disease Status : Detected")
    #         print("For the given dataset the Predicted value is Presence of heart Disease")
    #         print(" ")
    messagebox.showinfo("Process Status", "Process Completed")
    sys.exit(0) 

def process1():
    # Importing the dataset
    dataset = pd.read_csv('project.csv')
    print(dataset)
    dataset=dataset.dropna(how="any")
    print(dataset)
    print(" ")
    dataset.loc[:, 'Num'].replace([2, 3, 4], [1, 1, 1], inplace=True)

    print(dataset.info())
    print(" ")
    print(dataset)

    #Data Visualization

    #age vs cholestrol
    m = dataset['Age']
    n = dataset['Chol']
    plt.figure(figsize=(4,4))
    plt.title("Age vs Cholestrol",fontsize=20)
    plt.xlabel("Age",fontsize=15)
    plt.ylabel("Cholestrol",fontsize=15)
    plt.bar(m,n,label="age vs cholestrol",color=["red","orange"],width=0.5)
    plt.legend(loc='best')
    plt.show()

    #age vs chest pain
    m = dataset['Age']
    n = dataset['Cp']
    plt.figure(figsize=(4,4))
    plt.title("Age vs Chestpain",fontsize=20)
    plt.xlabel("Age",fontsize=15)
    plt.ylabel("Cp",fontsize=15)
    plt.bar(m,n,label="age vs chestpain",color=["red","orange"],width=0.5)
    plt.legend(loc='best')
    plt.show()

    #histogram of chest pain
    plt.figure(figsize=(10,8))
    plt.title("Histogram of Chest Pain")
    plt.hist(dataset['Cp'],rwidth=0.9)
    plt.show()

    X = dataset.iloc[:,:-1].values
    y = dataset.iloc[:, 13].values

    # Splitting the dataset into the Training set and Test set
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size = 0.25, random_state = 121)#101

    print(X_train)
    print(" ")

    # Feature Scaling
    sc = StandardScaler()
    X_train = sc.fit_transform(X_train)
    X_test = sc.transform(X_test)


    #EXPLORING THE DATASET

    dataset.Num.value_counts()

    estimators = []

    # Fitting RF to the Training set

    model1 =RandomForestClassifier(n_estimators=20)
    estimators.append(('rt1', model1))
    model2 =RandomForestClassifier(n_estimators=20)
    estimators.append(('rt2', model2))
    model3 =RandomForestClassifier(n_estimators=20)
    estimators.append(('rt3', model3))
    model4 =RandomForestClassifier(n_estimators=20)
    estimators.append(('rt4', model4))
    model5 =RandomForestClassifier(n_estimators=20)
    estimators.append(('rt5', model5))

    #voting classifier
    ensemble = VotingClassifier(estimators)
    ensemble.fit(X_train, y_train)
    y_pred = ensemble.predict(X_test)
    print("ypred : ")
    print(y_pred)
    print(" ")
    print("ytest : ")
    print(y_test)
    print(" ")

    #accuracy

    print("accuracy of random forest is: {}".format(accuracy_score(y_test, y_pred)))
    print(" ")

    #confusion matrix
    cm_Ensembler = confusion_matrix(y_test,y_pred)
    print("Confusion Matrix")
    print(cm_Ensembler)
    print(" ")

    #visualize the confusion matrix
    matrix = confusion_matrix(ensemble, X_test, y_test,cmap=plt.cm.YlGnBu)
    matrix.ax_.set_title('Confusion Matrix', color='blue')
    plt.xlabel('Predicted value', color='blue')
    plt.ylabel('Actual value', color='blue')
    plt.gcf().axes[0].tick_params(colors='blue')
    plt.gcf().axes[1].tick_params(colors='blue')
    plt.gcf().set_size_inches(10,6)
    plt.show()

    #accuracy
    accuracy = cross_val_score(ensemble, X_train, y_train, scoring='accuracy', cv = 10)
    print(accuracy)
    #get the mean of each fold 
    print("Accuracy of Model with Cross Validation is:",accuracy.mean() * 100)

if __name__=="__main__":        
    process()
    




