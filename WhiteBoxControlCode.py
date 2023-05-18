"""
Name:Francis Santos
Start date: 15/05/2023
Finish date: 22/05/2023
Title: Control code
"""

import pandas as pd
import re

def countertest(counter):
    try:
        #check which file the user is entering based on the parameter value passed in
        if counter == 0:
            file_loc = r'C:\Users\FrancisS\Downloads\AnonymisedSPlusData.xlsx'
            #input(r"Please enter pathway of First excel file here i.e., (C:\Users\JohnDoe\SyllabusPlusfile.xlsx) or type default if pathway has already been set for the syllabus plus output file:")
        if counter == 1:
            file_loc = r'C:\Users\FrancisS\Downloads\Copyofanonymised_names1.xlsx'
            #input(r"Please enter pathway of Second excel file here i.e., (C:\Users\JohnDoe\ContractHoursfile.xlsx) or type default if pathway has already been set for the lecturer contract hours file:")
        print('\n')
    #catch FileNotFoundErrors and prompt the user instead with a message
    except FileNotFoundError:
        print("Error! Incorrect pathway or pathway not found please try again")
    
    return file_loc

def testdefault():
    file_loc = 'default'
    xlfile = open("sampletest.txt","r")
    readfile = xlfile.readlines()
    if len(readfile) == 1 :
        file_loc = readfile[0].rstrip()
    xlfile.close()
    return file_loc

def file_sort(file_loc):
    #reads in excel file and places data in appropriate columns and formatted to more usable data
    try:
        read_data = pd.read_excel(file_loc,usecols="A,J,K,N,Q,U,V", names=["Module Name","Scheduled Start Time","Duration","Availability","Staff Names","Teaching Week Pattern","Number Of Teaching Weeks"])
    except FileNotFoundError:
        print("Error occurred retrieving file path, application terminating!")
        sys.exit()
    return True

def tesprocessF1(dataframe):
    if ((re.search(r"Weeks\s+([4-9]|1[0-6])\b", str(dataframe)))
            or (str(dataframe) == '0' and (re.search(r"\b([4-9]|1[0-6])\b", str(weekpattern))))
            or "Term 1" in str(dataframe)
             or (re.search(r"Week\s+([4-9]|1[0-6])\b", str(dataframe)))):
        return True
    else:
        result = "No pattern matches"
        return result
    
def tesprocessF1_2(wks_start_time, nOf_teaching_wks):
    if '00:00:00' in str(wks_start_time) and nOf_teaching_wks >= 13:
        result = 'weeks day time hours'
    if '00:00:00' not in str(wks_start_time) and nOf_teaching_wks >= 13:
        result = 'weeks unscheduled hours'           
    return result   

def tesprocesscalculationF1_3(wks_start_time,wks_sched_end,wks_sched_start):
    night_time = 18.0
    night_factor = 0.25
    if wks_start_time >= night_time:
        wks_sum1 = wks_sched_end - wks_sched_start
        wks_nighthrs = wks_sum1 * night_factor
    return wks_nighthrs

def tesprocesscalculationF1_4(wks_start_time,wks_sched_end,wks_sched_start):
    night_time = 8.0
    night_factor = 0.25
    if wks_start_time >= night_time:
        sumofhrs = wks_start_time + wks_sched_end + wks_sched_start
    return sumofhrs

def tesprocessF2(dataframe):
    if ((re.search(r"Weeks\s+(1[8-9]|2[0-9]|3[0-9]|4[0-5])\b", str(dataframe)))
            or (str(dataframe) == '0' and (re.search(r"1[8-9]|2[0-9]|3[0-9]4[0-5]", str(weekpattern))))
            or "Term 2" in str(dataframe)
            or "Term 3" in str(dataframe)
             or (re.search(r"Week\s+(1[8-9]|2[0-9]|3[0-9]|4[0-5])\b", str(dataframe)))):
        return True
    else:
        result = "No pattern matches"
        return result


    
def main():
    dataframe = 'Weeks 4-16'
    dataframe2 = 'Weeks 18-28'
    term = 'Term 1'
    wks_start_time = '00:00:00'
    nOf_teaching_wks = 13
    counter = 0
    file_path1 = countertest(counter)
    counter = 1
    file_path2 = countertest(counter)
    file_path = testdefault()
    a = file_sort(file_path1)
    p1result = tesprocessF1(dataframe)
    p1result2 = tesprocessF1(term)
    p1result3 = tesprocessF1_2(wks_start_time, nOf_teaching_wks)
    p1result4 = tesprocesscalculationF1_3(19.0,19.0,16)
    p1result5 = tesprocesscalculationF1_4(9.0,6.0,7.0)
    p2result = tesprocessF2(dataframe2)
if __name__== "__main__":
    main()
