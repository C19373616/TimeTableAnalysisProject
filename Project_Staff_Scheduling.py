"""
Project Name: Staff Scehduling
Creator: Francis Santos
Student Number: C19373616
Version: TAPv1.1 
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
    This function is used to read the file path the user has entered,
    then the excel file is read, the relevant columns are extracted, placed
    into a dataframe and columns are altered in preparation for math algorithms
    in another function. The column staff names were parsed in increments of 2
    and zipped into a tuple so that data pairs stored cannot be altered.The name
    pair list was defined as the new dataframe column and data was unstacked using
    the Python pandas module .explode() the sorted data was then passed back to main
    function.
    """
    read_data = pd.read_excel(file_loc,usecols="A,K,N,Q", names=["Module Name","Duration","Availability","Staff Names"])
    df = pd.DataFrame(read_data)
    df['Duration'] = df['Duration'].str.replace(r'\W','.',regex=True)
    df['Duration'] = df['Duration'].astype(float)
    names = df["Staff Names"].str.split(',')
    lst = []
    for i in range(0,len(names)):
        namepairs = list(zip(names[i][::2],names[i][1::2]))
        lst.append(namepairs)
    df["Staff Names"] = lst
    df = df.explode('Staff Names').reset_index(drop=True)
    df["Staff Names"] = df["Staff Names"].astype(str).replace(r"['()]","",regex=True)
    #df[['Last Name','First Name']] = df["Staff Names"].str.split(',',n=1,expand=True)       #n param stands for the number of splits done
    df.index += 1   
    df_order = df[["Staff Names","Duration","Availability","Module Name"]]
    return df_order
        

def main():
    set_configs()
    counter = 0
    while True:
        try:
            file_loc = input(r"Please enter pathway of the excel file here i.e., (C:\Users\JohnDoe\Data.xlsx) or type default if pathway has already been set:")
            xlfile = open("timetablelocation.txt","r")
            readfile = xlfile.readlines()
            xlfile.close()
            if len(readfile) <= 0:
                store_def_loc = input("Would you like to make this the default file location? Yes or No")
                if "yes" in store_def_loc.lower():
                    save_loc = open("timetablelocation.txt","w")
                    save_loc.write(file_loc)
                    save_loc.close()
                xlfile.close()
            elif "default" in file_loc.lower():
                if len(readfile) > 1:
                    for i in readfile:
                        counter += 1
                        print(counter," - ",i)
                    whichloc = input("More than 1 default location detected, please specify which file location to use 1 or 2 etc. : ")
                    file_loc = readfile
                    sys.exit()
                elif len(readfile) == 1 :
                    file_loc = readfile
                else:
                    print("No default file location found or set")
                xlfile.close()
            dataframe = file_sort(file_loc)
            if len(dataframe) > 0:
                break
        except FileNotFoundError:
            print("Error! Incorrect pathway please try again")
    for i in range(1,len(dataframe["Duration"])):
        if "Angela" in dataframe["First Name"][i] and "Adams" in dataframe["Last Name"][i] and "Semester 2"  in dataframe["Availability"][i]:
            print(dataframe["Duration"][i],i)
           
    print("\n",dataframe.head(600))
    

if __name__== "__main__":
    main()
    
"""
Test Code:
#df[['Last Name','First Name']] = df["Staff Names"].str.split(',',n=1,expand=True)   #n param stands for the number of splits done
#df["Staff Names"] = df["Staff Names"].str.split(',').str[:2].str.join(',')
#print(df['Staff Names'].head(600)) 
"""
