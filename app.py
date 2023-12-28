import pandas as pd
import numpy as np


# Introduction
# This is a script automating the anual processing of K.C.P.E results,
# It was ment for State House Primary School 
# It was writen in 2020, to be used in the years: 2020,2021,2022, 2023 until the change of the education system
# In the event of changes in the Governments system changes 

# For this script to work it needs 2 csv files located in the same directory(folder as the script): 
# 1: Candidates Data
# A file named kcpe_can.csv containing the candidates data, 
# A file contains the following columns:
# Index No. : Candidates index number
# NAME : Candidates names
# Class : Candidates name

# 2: Candidates KCPE results downloaded from the KNEC portal 
# A file named kcpe_results_n downloaded from the KNEC kcpe portal


#Get the current candidates Data containing their class, index number, 
current_class= pd.read_csv('kcpe_can.csv')
#Get the results data to be processed
results= pd.read_csv('kcpe_results_n.csv')









#Get Categorisation variables for the new processed data
#The indexes and names
candidates=results['SCHOOL_NO'].unique()

#All the subjects they did
subjects=results['Textbox20'].unique()

#All the classes
classes=current_class['Class'].unique()





#Defining the new Data
columns=['Index_No', 'Name', 'Sex', 'Class']
#Adding subjects
for subject in subjects:
    columns.append(subject)
columns.append('Marks')
#Creating a dataframe
processed_results=pd.DataFrame(columns=columns)
#Defining the index
indexes=[]
for index in candidates:
    indexes.append(index.split(' ')[0])
processed_results['Index_No']=indexes
processed_results=processed_results.set_index('Index_No')




#A function that gets the class of the candidate from current data
def get_class(class_data,cand_index):
    for i in class_data.index:
        if(class_data.iloc[i,0]== int(cand_index)):
            return class_data.iloc[i,2]
    return "Not found"



#The actual processing
for i in results.index:
    can_index=results.iloc[i,0:]['SCHOOL_NO'].split(' ')[0]
    name=results.iloc[i,0:]['SCHOOL_NO'].split('     ')[1]
    sex=results.iloc[i,0:]['SEX'].split(' ')[0]
    subject=results.iloc[i,0:]['Textbox20']
    total_mark=results.iloc[i,0:]['Textbox41']
    subject_mark=results.iloc[i,0:]['MKS']
    total_mark=results.iloc[i,0:]['Textbox41']
    #Assigning values
    processed_results.loc[can_index,'Name']=name
    processed_results.loc[can_index,'Sex']=sex
    processed_results.loc[can_index,'Marks']=total_mark
    processed_results.loc[can_index,subject]=subject_mark
    processed_results.loc[can_index,subject]=subject_mark
    processed_results.loc[can_index,'Class']=get_class(current_class,can_index)

processed_results.sort_values(by=['Marks'])



#Filtering the needed data
#The last of the List
processed_results=processed_results.sort_values(by=['Marks'])
bottom_15=processed_results.head(15)
bottom_15

#The top of the list
processed_results=processed_results.sort_values(by=['Marks'], ascending=False)
top_15=processed_results.head(15)
top_15

#Females
female=processed_results[processed_results['Sex']=='*F*']

#Males
male=processed_results[processed_results['Sex']=='*M*']

#Separating according to individual classes
list_of_class_results={}
for Class in classes:
    class_result=processed_results[processed_results['Class']==Class]
    list_of_class_results[Class]=class_result



#Storing all processed information in excell file for use
with pd.ExcelWriter('output.xlsx') as writer:  
    processed_results.to_excel(writer, sheet_name='KCPE 2022')
    top_15.to_excel(writer, sheet_name='Top 15')
    bottom_15.to_excel(writer, sheet_name='Bottom 15')
    male.to_excel(writer, sheet_name='Male')
    female.to_excel(writer, sheet_name='Female')
    for class_result in list_of_class_results:
        list_of_class_results[class_result].to_excel(writer, sheet_name=class_result)