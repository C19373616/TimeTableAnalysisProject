"""
Project Name: Staff Scehduling
Creator: Francis Santos
Student Number: C19373616
Version: TAPv2.3  
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
    """
    file_setup():
    This function is used to set the configurations for pandas dataframe to allow for the correct
    display to be shown when the program is executed, the max number of rows and columns are set
    as well as a print to explain what the program is.
    """
    pd.set_option('display.max_rows', 1000)
    pd.set_option('display.max_columns', 1000)
    pd.set_option('max_colwidth', None)
    pd.set_option('display.width', 1000)
    print("!========================================================================!")
    print("Hello, this program is used to summarise the hours for TUDublin staff")
    print("Please ensure that file is a .xlsx and include .xlsx in the full pathway")
    print("!========================================================================!")

def file_setup():
    """
    file_setup():
    This function is used to define the path of the excel file that the data will be extracted from.
    In this function, if there are no previous file paths defined in the .txt file then the user
    will be asked if they want to make the file path they entered when they run the program the default
    file location. This way the user does not need to constantly enter the file path of the excel file.
    The file location will be passed to the file_sort() function if this function detects that the path
    is incorrect the application will terminate.
    """
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
                    whichloc = int(input("More than 1 default location detected, please specify which file location to use 0, 1 or 2 etc. : "))
                    a = readfile[whichloc].rstrip()
                    file_loc = a
            elif len(readfile) == 1 :
                    file_loc = readfile
            else:
                print("No default file location found or set")
            xlfile.close()
            if len(file_loc) > 0:
                break
            else:
                print("Error occurred retrieving file path, application terminating!")
                sys.exit()
        except FileNotFoundError:
            print("Error! Incorrect pathway or pathway not found please try again")
    return file_loc

def file_sort(file_loc):
    """
    file_sort():
    This function is used to read the excel file stored in the file path defined, from the excel file
    the relevant columns are extracted, placedinto a dataframe and columns are altered in preparation
    for math algorithms in another function. The column staff names were parsed in increments of 2
    and zipped into a tuple so that data pairs stored cannot be altered. The name pair list was defined
    as the new dataframe column and data was unstacked using the Python pandas module .explode() and
    dataframe was re-indexed. Then using regex negative lookbehind and negative lookahead it looks for
    single quotes that is not preceded by a word or not followed by a word and also removes the set of
    round brackets. Finally the columns are arranged in the order distinguished, then the .unique()
    module,unique values are taken from the column staff names to get a list of non-duplicated names
    preparing it for staff total hours in the process_data() function. The unique list and ordered
    dataframe are sent back to main.
    """
    try:
        read_data = pd.read_excel(file_loc,usecols="A,J,K,N,Q,U,V", names=["Module Name","Scheduled Start Time","Duration","Availability","Staff Names","Teaching Week Pattern","Number Of Teaching Weeks"])
    except FileNotFoundError:
        print("Error occurred retrieving file path, application terminating!")
        sys.exit()
    df = pd.DataFrame(read_data)
    df.fillna(0,inplace=True)
    df['Scheduled Start Time'] = df['Scheduled Start Time'].replace(0,'00:00:00')
    df['Scheduled Start Time'] = pd.to_datetime(df['Scheduled Start Time'], format='%H:%M:%S')
    df['Duration'] = pd.to_datetime(df['Duration'], format='%H:%M')
    names = df["Staff Names"].str.split(',')
    lst = []
    for i in range(0,len(names)):
        namepairs = list(zip(names[i][::2],names[i][1::2]))
        lst.append(namepairs)
    df["Staff Names"] = lst
    df = df.explode('Staff Names').reset_index(drop=True)
    df["Staff Names"] = df["Staff Names"].astype(str).replace(r"(?<!\w)'|'(?!\w)|[()]",'',regex=True)
    df["Staff Names"] = df["Staff Names"].astype(str).replace(r'[""]','',regex=True)
    df.index += 1
    df_order = df[["Staff Names","Scheduled Start Time","Duration","Availability","Module Name","Teaching Week Pattern","Number Of Teaching Weeks"]]
    uniqlst = df["Staff Names"].unique()
    uniqlst.sort()
    return df_order,uniqlst

def process_sem1_data(dataframe, uniqlst):
    sem1_lst = []
    #formats the dataframe to allow for easier data parsing
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"(-)",',',regex=True)
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"\(",'',regex=True)
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"\)",'',regex=True)
    #use of timedelta reference: https://www.geeksforgeeks.org/python-datetime-timedelta-class/
    for a in range(0, len(uniqlst)):
        night_time = pd.Timedelta(hours=18, minutes=0, seconds=0)
        night_factor = 0.25
        unsched_hrs = pd.Timedelta(0)
        control = pd.Timedelta(0)
        counter = pd.Timedelta(0)
        counter1 = pd.Timedelta(0)
        nightcount = pd.Timedelta(0)
        totalhrs = pd.Timedelta(0)
        wks_counter = pd.Timedelta(0)
        wks_counter1 = pd.Timedelta(0)
        wks_nightcount = pd.Timedelta(0)
        wks_totalhrs = pd.Timedelta(0)
        wks_realhrs = pd.Timedelta(0)
        for i in range(1, len(dataframe["Staff Names"])):
            #\b is a word boundary which essentially allows only position between the boundary defined to be matched meaning if anything follows the number it is not matched.
            #reference: https://medium.com/factory-mind/regex-tutorial-a-simple-cheatsheet-by-examples-649dc1c3f285
            #if statement finds semester, weeks and terms all related to semester 1
            if (uniqlst[a] in dataframe["Staff Names"][i] and
                    (("Semester 1" in str(dataframe["Availability"][i]) or "Term 1" in str(dataframe["Availability"][i]))
                     or (str(dataframe["Availability"][i]) == '0')
                     or (re.search(r"Weeks\s+([4-9]|1[0-6])\b", str(dataframe["Availability"][i])))
                     or (re.search(r"Week\s+([4-9]|1[0-6])\b", str(dataframe["Availability"][i]))))):
                #section uses regex to find weeks from 4 - 9 and 10 - 16 
                if ((re.search(r"Weeks\s+([4-9]|1[0-6])\b", str(dataframe["Availability"][i])))
                    or (str(dataframe["Availability"][i]) == '0' and (re.search(r"([\b[4-9]\b|1[0-6]\b)", str(dataframe["Teaching Week Pattern"][i]))))
                    or "Term 1" in str(dataframe["Availability"][i])
                     or (re.search(r"Week\s+([4-9]|1[0-6])\b", str(dataframe["Availability"][i])))):
                    nOf_teaching_wks = int(dataframe["Number Of Teaching Weeks"][i])
                    #index_Ahead = int(dataframe["Number Of Teaching Weeks"][i])
                    wks_counter += pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    wks_counter1 = pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    wks_start_time = pd.Timestamp(dataframe["Scheduled Start Time"][i])
                    wks_sched_start = pd.Timedelta(hours=wks_start_time.hour, minutes=wks_start_time.minute)
                    wks_sched_end = wks_counter1 + wks_sched_start
                    wks_totalhrs = wks_counter
                    wks_convrt_13 = 0
                    wks_nighthrs = pd.Timedelta(0)
                    if wks_sched_end > night_time:
                        if wks_sched_start >= night_time:
                            wks_sum1 = wks_sched_end - wks_sched_start
                            wks_nighthrs = wks_sum1 * night_factor
                        elif wks_sched_start < night_time:
                            wks_sum2 = wks_sched_end - night_time
                            wks_nighthrs = wks_sum2 * night_factor
                        wks_nightcount += wks_nighthrs
                    if wks_nightcount != control:
                        wks_totalhrs = wks_nightcount + wks_counter
                    if nOf_teaching_wks != control and nOf_teaching_wks < 13:
                        #print(wks_counter1,"/13",nOf_teaching_wks)
                        wks_convrt_13 = wks_counter1 * (nOf_teaching_wks/13)
                        if wks_nightcount != control:
                            wks_convrt_13_night = wks_nighthrs * (nOf_teaching_wks/13)
                            wks_convrt_13 = wks_convrt_13 + wks_convrt_13_night
                        """
                        HERE!!
                        if '00:00:00' in str(wks_start_time):
                            print(dataframe["Staff Names"][i],unsched_hrs,"True")
                            unsched_hrs += wks_convrt_13
                        """
                        #print(wks_convrt_13,nOf_teaching_wks)
                        wks_realhrs += wks_convrt_13
 
                else:
                    counter += pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    counter1 = pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    nOf_teaching_wks = int(dataframe["Number Of Teaching Weeks"][i])
                    start_time = pd.Timestamp(dataframe["Scheduled Start Time"][i])
                    sched_start = pd.Timedelta(hours=start_time.hour, minutes=start_time.minute)
                    sched_end = counter1 + sched_start
                    totalhrs = counter
                    nighthrs = pd.Timedelta(0)
                    if sched_start == pd.Timedelta(0) and nOf_teaching_wks >= 13:
                        unsched_hrs += counter1
                        #print(dataframe["Staff Names"][i],unsched_hrs,"True") 
                    if sched_end > night_time:
                        if sched_start >= night_time:
                            sum1 = sched_end - sched_start
                            nighthrs = sum1 * night_factor
                            #print(nighthrs,uniqlst[a])
                        elif sched_start < night_time:
                            sum2 = sched_end - night_time
                            nighthrs = sum2 * night_factor
                        nightcount += nighthrs
                        #print(nightcount)
                    #print(totalhrs,wks_realhrs)
            if nightcount != control:
                totalhrs = nightcount + counter
        if wks_totalhrs != control:
            #print(totalhrs,wks_realhrs)
            totalhrs = totalhrs + wks_realhrs
        totalseconds = totalhrs.total_seconds()
        daysconvert_f = int(totalseconds/86400)*24
        hoursconver_f = (totalseconds%86400)/3600
        totalhrs_asfloat = daysconvert_f + hoursconver_f
        totalhrs_asfloat = round(totalhrs_asfloat,2)
        #print(totalhrs_asfloat, uniqlst[a])
        sem1_lst.append(totalhrs_asfloat)
    return sem1_lst

def process_sem2_data(dataframe, uniqlst):
    sem2_lst = []
    #formats the dataframe to allow for easier data parsing
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"(-)",',',regex=True)
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"\(",'',regex=True)
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"\)",'',regex=True)
    #use of timedelta reference: https://www.geeksforgeeks.org/python-datetime-timedelta-class/
    for a in range(0, len(uniqlst)):
        night_time = pd.Timedelta(hours=18, minutes=0, seconds=0)
        night_factor = 0.25
        control = pd.Timedelta(0)
        counter = pd.Timedelta(0)
        counter1 = pd.Timedelta(0)
        nightcount = pd.Timedelta(0)
        totalhrs = pd.Timedelta(0)
        wks_counter = pd.Timedelta(0)
        wks_counter1 = pd.Timedelta(0)
        wks_nightcount = pd.Timedelta(0)
        wks_totalhrs = pd.Timedelta(0)
        wks_realhrs = pd.Timedelta(0)
        for i in range(1, len(dataframe["Staff Names"])):
            #\b is a word boundary which essentially allows only position between the boundary defined to be matched meaning if anything follows the number it is not matched.
            #reference: https://medium.com/factory-mind/regex-tutorial-a-simple-cheatsheet-by-examples-649dc1c3f285
            #if statement finds semester, weeks and terms all related to semester 1
            if (uniqlst[a] in dataframe["Staff Names"][i] and
                    (("Semester 2" in str(dataframe["Availability"][i]) or "Term 2" in str(dataframe["Availability"][i]))
                     or "Semester 1&2" in str(dataframe["Availability"][i]).lstrip()
                     or "Term 3" in str(dataframe["Availability"][i])
                     or (str(dataframe["Availability"][i]) == '0')
                     or (re.search(r"(1[8-9]|2[0-9]|3[0-9]4[0-5])", str(dataframe["Availability"][i])))
                     or (re.search(r"Weeks\s+(1[8-9]|2[0-9]|3[0-9]|4[0-5])\b", str(dataframe["Availability"][i])))
                     or (re.search(r"Week\s+(1[8-9]|2[0-9]|3[0-9]|4[0-5])\b", str(dataframe["Availability"][i]))))):
                #section uses regex to find weeks from 18 - 19, 22 - 29 and 30 - 39
                if ((re.search(r"Weeks\s+(1[8-9]|2[0-9]|3[0-9]|4[0-5])\b", str(dataframe["Availability"][i])))
                    or (str(dataframe["Availability"][i]) == '0' and (re.search(r"(1[8-9]|2[0-9]|3[0-9])", str(dataframe["Teaching Week Pattern"][i]))))
                    or "Term 2" in str(dataframe["Availability"][i])
                     or "Term 3" in str(dataframe["Availability"][i])
                     or (re.search(r"Week\s+(1[8-9]|2[0-9]|3[0-9]|4[0-5])\b", str(dataframe["Availability"][i])))):
                    nOf_teaching_wks = int(dataframe["Number Of Teaching Weeks"][i])
                    wks_counter += pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    wks_counter1 = pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    wks_start_time = pd.Timestamp(dataframe["Scheduled Start Time"][i])
                    wks_sched_start = pd.Timedelta(hours=wks_start_time.hour, minutes=wks_start_time.minute)
                    wks_sched_end = wks_counter1 + wks_sched_start
                    wks_totalhrs = wks_counter
                    wks_convrt_13 = 0
                    wks_nighthrs = pd.Timedelta(0)
                    if wks_sched_end > night_time:
                        if wks_sched_start >= night_time:
                            wks_sum1 = wks_sched_end - wks_sched_start
                            wks_nighthrs = wks_sum1 * night_factor
                        elif wks_sched_start < night_time:
                            wks_sum2 = wks_sched_end - night_time
                            wks_nighthrs = wks_sum2 * night_factor
                        wks_nightcount += wks_nighthrs
                    if wks_nightcount != control:
                        wks_totalhrs = wks_nightcount + wks_counter
                        #print(wks_sched_end)                        
                    if nOf_teaching_wks != control and nOf_teaching_wks < 13:
                        wks_convrt_13 = wks_counter1 * (nOf_teaching_wks/13)
                        #print(wks_counter1,"/13")
                        if wks_nightcount != control:
                            wks_convrt_13_night = wks_nighthrs * (nOf_teaching_wks/13)
                            #print("night count",wks_convrt_13_night)
                            wks_convrt_13 = wks_convrt_13 + wks_convrt_13_night
                        #print(wks_convrt_13,nOf_teaching_wks)
                        wks_realhrs += wks_convrt_13
                            
                else:
                    counter += pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    counter1 = pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    start_time = pd.Timestamp(dataframe["Scheduled Start Time"][i])
                    nOf_teaching_wks = int(dataframe["Number Of Teaching Weeks"][i])
                    sched_start = pd.Timedelta(hours=start_time.hour, minutes=start_time.minute)
                    sched_end = counter1 + sched_start
                    totalhrs = counter
                    nighthrs = pd.Timedelta(0)
                    if sched_end > night_time:
                        if sched_start >= night_time:
                            sum1 = sched_end - sched_start
                            nighthrs = sum1 * night_factor
                            #print(nighthrs,uniqlst[a])
                        elif sched_start < night_time:
                            sum2 = sched_end - night_time
                            nighthrs = sum2 * night_factor
                        nightcount += nighthrs
                        #print(nightcount)
            if nightcount != control:
                totalhrs = nightcount + counter
        if wks_totalhrs != control:
            #print(totalhrs, wks_realhrs)
            totalhrs = totalhrs + wks_realhrs
        totalseconds = totalhrs.total_seconds()
        daysconvert_f = int(totalseconds/86400)*24
        hoursconver_f = (totalseconds%86400)/3600
        totalhrs_asfloat = daysconvert_f + hoursconver_f
        totalhrs_asfloat = round(totalhrs_asfloat,2)
        #print(totalhrs_asfloat , uniqlst[a])
        sem2_lst.append(totalhrs_asfloat)
    return sem2_lst

def data_analysis(sem1_lst,sem2_lst,uniqlst):
    usrinpt = input(r'Please enter file location of the contract hours file i.e., (C:\Users\JohnDoe\Data.xlsx): ')
    xl = pd.ExcelFile(usrinpt)
    read1 = pd.read_excel(xl, 'Lecturers',usecols="A,B,C", names=["Lecturers1","S1 Hours","S2 Hours"])
    contracth_df = pd.DataFrame(read1)
    realtime_df = pd.DataFrame(sem1_lst,columns=["S1 Hours"])
    #use of strftime reference: https://www.programiz.com/python-programming/datetime/strftime
    #realtime_df.index += 1
    #realtime_df['S1 Hours'] = pd.to_datetime(realtime_df['S1 Hours'].dt.total_seconds(), unit='s').dt.strftime('%H:%M:%S')
    realtime_df['S2 Hours'] = sem2_lst
    realtime_df['Lecturers'] = uniqlst
    realtime_df = realtime_df[['Lecturers','S1 Hours','S2 Hours']]
    realtime_df['Lecturers'] = realtime_df['Lecturers'].astype(str).replace(r"(  )",'',regex=True)
    contracth_df['Lecturers1'] = contracth_df['Lecturers1'].astype(str).replace(r"( )",'',regex=True)
    staffhrstotalhrsS1 = []
    staffhrstotalhrsS2 = []
    staffhrsS1_undr_over = []
    staffhrsS2_undr_over = []
    total_under_over = []
    yearhrs = []
    customer_outputRep = pd.DataFrame()
    print("\n")
    for i in range(0,len(contracth_df['Lecturers1'])):
        if contracth_df['Lecturers1'][i] in str(realtime_df['Lecturers']).lstrip():
            #print(contracth_df['Lecturers1'][i],contracth_df['S1 Hours'][i],realtime_df['S1 Hours'][i])
            indx = realtime_df.index[realtime_df['Lecturers'] == contracth_df['Lecturers1'][i]][0]
            staffhrstotalhrsS1.append(realtime_df['S1 Hours'][indx])
            staffhrstotalhrsS2.append(realtime_df['S2 Hours'][indx])
        else:
            staffhrstotalhrsS1.append(0)
            staffhrstotalhrsS2.append(0)
    for i in range(0,len(contracth_df['Lecturers1'])):
        s1_chs_sub_ttlhrs = float(staffhrstotalhrsS1[i]) - float(contracth_df['S1 Hours'][i])
        s2_chs_sub_ttlhrs = float(staffhrstotalhrsS2[i]) - float(contracth_df['S2 Hours'][i])
        yeartotal = (float(staffhrstotalhrsS1[i])*13)+(float(staffhrstotalhrsS2[i])*13)
        staffhrsS1_undr_over.append(s1_chs_sub_ttlhrs)
        staffhrsS2_undr_over.append(s2_chs_sub_ttlhrs)
        yearhrs.append(yeartotal)
    for i in range(0,len(contracth_df['Lecturers1'])):
        totalundrover = float(staffhrsS1_undr_over[i]) + float(staffhrsS2_undr_over[i])
        
        total_under_over.append(totalundrover)
        
    customer_outputRep['Lecturers'] = contracth_df['Lecturers1']
    customer_outputRep['CHS1'] = contracth_df['S1 Hours']
    customer_outputRep['S1 Total Hours'] = staffhrstotalhrsS1
    customer_outputRep['S1 Over'] = staffhrsS1_undr_over
    
    customer_outputRep['CHS2'] = contracth_df['S2 Hours']
    customer_outputRep['S2 Total Hours'] = staffhrstotalhrsS2
    customer_outputRep['S2 Over'] = staffhrsS2_undr_over
    customer_outputRep['Over Hrs'] = total_under_over
    customer_outputRep['Year'] = yearhrs
    customer_outputRepA = customer_outputRep[['Lecturers','CHS1','S1 Over','CHS2','S2 Over','Over Hrs','Year']]
    print(customer_outputRepA )
    #print(realtime_df)
      

def main():
    set_configs()
    file_path = file_setup()
    dataframe, uniqlst = file_sort(file_path)
    if len(dataframe) > 0 or len(uniqlst) > 0:
        sem1_lst = process_sem1_data(dataframe, uniqlst)
        print("\n")
        sem2_lst = process_sem2_data(dataframe, uniqlst)
    else:
        print("Error occurred retrieving data application terminating")
        sys.exit()
    data_analysis(sem1_lst,sem2_lst,uniqlst)    
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

            if len(dataframe) > 0 and len(uniqlst) > 0:
                #print(len(file_loc))
                break
            else:
                print("Error occurred retrieving data application terminating")
                sys.exit()

"""
