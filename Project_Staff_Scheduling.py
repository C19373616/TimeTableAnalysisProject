"""
Project Name: Staff Scehduling
Creator: Francis Santos
Student Number: C19373616
Version: TAPv1.2  
"""

"""
- The section below imports necessary necessary libraries and
correct parameters are set relating to appropriate data display
- Openpyxl is a library used to read or write microsoft excel files 
"""
import pandas as pd
import re
import sys
import numpy as np
import openpyxl

def set_configs():   
    pd.set_option('display.max_rows', 1000)
    pd.set_option('display.max_columns', 1000)
    pd.set_option('max_colwidth', None)
    pd.set_option('display.width', 1000)
    print("!========================================================================!")
    print("Hello, this program is used to summarise the hours for TUD staff")
    print("Please ensure that file is a .xlsx and include .xlsx in the full pathway")
    print("!========================================================================!")


def file_sort(file_loc):
    """
    file_sort():
    This function is used to read the file path the user has entered,then the excel file is read,
    the relevant columns are extracted, placedinto a dataframe and columns are altered in preparation
    for math algorithms in another function. The column staff names were parsed in increments of 2
    and zipped into a tuple so that data pairs stored cannot be altered. The name pair list was defined
    as the new dataframe column and data was unstacked using the Python pandas module .explode() and
    dataframe was re-indexed. Then using regex negative lookbehind and negative lookahead it looks for
    single quotes that is not preceded by a word or not followed by a word and also removes the set of
    round brackets.
    """
    read_data = pd.read_excel(file_loc,usecols="A,J,K,N,Q", names=["Module Name","Scheduled Start Time","Duration","Availability","Staff Names"])
    df = pd.DataFrame(read_data)
    df.fillna(0,inplace=True)
    df['Duration'] = df['Duration'].str.replace(r'\W','.',regex=True)
    df['Duration'] = df['Duration'].astype(float)
    names = df["Staff Names"].str.split(',')
    lst = []
    for i in range(0,len(names)):
        namepairs = list(zip(names[i][::2],names[i][1::2]))
        lst.append(namepairs)
    df["Staff Names"] = lst
    df = df.explode('Staff Names').reset_index(drop=True)
    df["Staff Names"] = df["Staff Names"].astype(str).replace(r"(?<!\w)'|'(?!\w)|[()]",'',regex=True)
    df["Staff Names"] = df["Staff Names"].astype(str).replace(r'[""]','',regex=True)    
    #df[['Last Name','First Name']] = df["Staff Names"].str.split(',',n=1,expand=True)       #n param stands for the number of splits done
    df.index += 1
    df_order = df[["Staff Names","Scheduled Start Time","Duration","Availability","Module Name"]]
    uniqlst = df["Staff Names"].unique()
    #xlsx = pd.ExcelFile(r'C:\Users\FrancisS\Downloads\Copyofanonymised_names1.xlsx')
    #df1 = pd.read_excel(xlsx, 'Lecturers',usecols="A,B,C", names=["Lecturers1","S1 Hours","S2 Hours"])
    uniqlst.sort()
    #print(len(unqlst))
       
    #a = df1.sort_values("Lecturers1")
    #print(a)
    #print(df.head(600))
    print(type(uniqlst))
    return df_order,uniqlst

def reformat(dataframe,uniqlst):
    for a in range(0,len(uniqlst)):
        counter = 0 
        for i in range(1,len(dataframe["Staff Names"])):
            if uniqlst[a] in dataframe["Staff Names"][i] and "Semester 1" in str(dataframe["Availability"][i]) :
                counter += dataframe["Duration"][i]
        print(counter,uniqlst[a])
        
        
               
        

def main():
    set_configs()
    while True:
        try:
            file_loc = input(r"Please enter pathway of the excel file here i.e., (C:\Users\JohnDoe\Data.xlsx) or type default if pathway has already been set:")
            #hardcode = r"C:\Users\FrancisS\Downloads\AnonymisedSPlusData.xlsx"
            xlfile = open("timetablelocation.txt","r")
            readfile = xlfile.readlines()
            if len(readfile) <= 0:
                store_def_loc = input("Would you like to make this the default file location? Yes or No")
                if "yes" in store_def_loc.lower():
                    save_loc = open("timetablelocation.txt","w")
                    save_loc.write(file_loc)
                    save_loc.close()
                xlfile.close()
            elif "default" in file_loc.lower():
                if len(readfile) > 1:
                    counter = 0
                    for i in readfile:
                        print(counter," - ",i)
                        counter += 1
                    whichloc = int(input("More than 1 default location detected, please specify which file location to use 1 or 2 etc. : "))
                    a = readfile[whichloc].rstrip()
                    file_loc = a
            elif len(readfile) == 1 :
                    file_loc = readfile
            else:
                print("No default file location found or set")
            xlfile.close()
            dataframe, uniqlst = file_sort(file_loc)
            if len(dataframe) > 0 and len(uniqlst) > 0:
                break
            else:
                print("Error occurred retrieving data application terminating")
                sys.exit()
        except FileNotFoundError:
            print("Error! Incorrect pathway or pathway not found please try again")

    reformat(dataframe, uniqlst)

    #print("\n",dataframe.head(600))
    

if __name__== "__main__":
    main()
    
"""
Test Code:
#df[['Last Name','First Name']] = df["Staff Names"].str.split(',',n=1,expand=True)   #n param stands for the number of splits done
#df["Staff Names"] = df["Staff Names"].str.split(',').str[:2].str.join(',')
#print(df['Staff Names'].head(600))

for i in range(1,len(dataframe["Duration"])):
#if "Angela" in dataframe["First Name"][i] and "Adams" in dataframe["Last Name"][i] and "Semester 2"  in dataframe["Availability"][i]:
if "Angela" in dataframe["Staff Names"][i] and "Adams" in dataframe["Staff Names"][i] and "Semester 1"  in dataframe["Availability"][i] :
print(i,dataframe["Duration"][i])
"""
