"""
Project Name: Staff Scehduling
Creator: Francis Santos
Student Number: C19373616
Version: TAPv1.4  
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
    print("Hello, this program is used to summarise the hours for TUDublin staff")
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
    #df['Duration'] = df['Duration'].str.replace(r'\W','.',regex=True)
    #df['Duration'] = df['Duration'].astype(float)
    df['Scheduled Start Time'] = df['Scheduled Start Time'].replace(0,'00:00:00')
    df['Scheduled Start Time'] = pd.to_datetime(df['Scheduled Start Time'], format='%H:%M:%S')
    df['Duration'] = pd.to_datetime(df['Duration'], format='%H:%M')
    #df['Duration'] = df['Duration'].dt.time
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
    #df1 = pd.read_excel(xlsx, 'Lecturers',usecols="A,B,C", names=["Lecturers1","S1 Hours","S2 Hours"])
    uniqlst.sort()
    #print(df["Duration"])
    return df_order,uniqlst

def reformat(dataframe, uniqlst):
    sem1weeks = ['4','5','6','7','8','9','10','11','12','13','14']
    sem1_lst = []
    sem2_lst = []
    #use of timedelta reference: https://www.geeksforgeeks.org/python-datetime-timedelta-class/
    for a in range(0, len(uniqlst)):
        night_time = pd.Timedelta(hours=18, minutes=0, seconds=0)
        night_factor = 0.25
        medium = pd.Timedelta(0)
        counter = pd.Timedelta(0)
        counter1 = pd.Timedelta(0)
        nightcount = pd.Timedelta(0)
        totalhrs = pd.Timedelta(0)
        for i in range(1, len(dataframe["Staff Names"])):
            if uniqlst[a] in dataframe["Staff Names"][i] and ("Semester 1" in str(dataframe["Availability"][i]) or "Term 1" in str(dataframe["Availability"][i])) :
                counter += pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                counter1 = pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                start_time = pd.Timestamp(dataframe["Scheduled Start Time"][i])
                sched_start = pd.Timedelta(hours=start_time.hour, minutes=start_time.minute)
                sched_end = counter1 + sched_start
                totalhrs = counter
                if sched_end > night_time:
                    if sched_start >= night_time:
                        sum1 = sched_end - sched_start
                        nighthrs = sum1 * night_factor
                    elif sched_start < night_time:
                        sum2 = sched_end - night_time
                        nighthrs = sum2 * night_factor
                    nightcount += nighthrs
            if nightcount != medium:
                totalhrs = nightcount + counter
        print(totalhrs, uniqlst[a])
        sem1_lst.append(totalhrs)
    xl = pd.ExcelFile(r'C:\Users\franc.LAPTOP-CMCLL6GJ\Downloads\anonymised_names1.xlsx')
    read1 = pd.read_excel(xl, 'Lecturers',usecols="A,B,C", names=["Lecturers1","S1 Hours","S2 Hours"])
    contracth_df = pd.DataFrame(read1)
    realtime_df = pd.DataFrame(sem1_lst,columns=["S1 Hours"])
    #use of strftime reference: https://www.programiz.com/python-programming/datetime/strftime
    #realtime_df.index += 1
    realtime_df['S1 Hours'] = pd.to_datetime(realtime_df['S1 Hours'].dt.total_seconds(), unit='s').dt.strftime('%H:%M:%S')
    realtime_df['Lecturers'] = uniqlst
    realtime_df['Lecturers'] = realtime_df['Lecturers'].astype(str).replace(r"(  )",'',regex=True)
    
    contracth_df['Lecturers1'] = contracth_df['Lecturers1'].astype(str).replace(r"( )",'',regex=True)
    lst = []
    print("\n")
    print(len(realtime_df['Lecturers']))
    print(contracth_df['Lecturers1'][3],realtime_df['Lecturers'][1].lstrip())
    conter = 0
    for i in range(0,len(contracth_df['Lecturers1'])):
        if contracth_df['Lecturers1'][i] in str(realtime_df['Lecturers']).lstrip():
            conter += 1
            print(conter, contracth_df['Lecturers1'][i])
        else:
            lst.append(contracth_df['Lecturers1'][i])
    print(lst)
    



            
               
        

def main():
    set_configs()
    while True:
        try:
            file_loc = input(r"Please enter pathway of the excel file here i.e., (C:\Users\JohnDoe\Data.xlsx) or type default if pathway has already been set:")
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
#print(len(unqlst))
       
#a = df1.sort_values("Lecturers1")
#print(a)
#print(df.head(600))
"""
