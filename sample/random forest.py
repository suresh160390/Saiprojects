import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestClassifier
from sklearn.ensemble import VotingClassifier
from sklearn.metrics import confusion_matrix
from sklearn.metrics import accuracy_score
from sklearn import model_selection
from sklearn.metrics import plot_confusion_matrix
from sklearn.model_selection import cross_val_score


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
    matrix = plot_confusion_matrix(ensemble, X_test, y_test,cmap=plt.cm.YlGnBu)
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
