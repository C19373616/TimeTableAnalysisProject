"""
Project Name: Staff Scheduling
Creator: Francis Santos
Student Number: C19373616
Version: TAPv3.1  
"""

"""
- The section below imports necessary necessary libraries and
correct parameters are set relating to appropriate data display
- Openpyxl is a library used to read or write microsoft excel files 
"""
import pandas as pd
import re
import sys
import openpyxl

def set_configs():
    """
    set_configs():
    This function is used to set the configurations for pandas dataframe to allow for the correct
    display to be shown when the program is executed, the max number of rows and columns are set
    as well as a print to explain what the program is.
    """
    #limit display of data 
    pd.set_option('display.max_rows', 500)
    pd.set_option('display.max_columns', None)
    pd.set_option('max_colwidth', 500)
    pd.set_option('display.width', 500)
    #display a guide for user to use
    print("!===================================================================================!")
    print("Hello, this program is used to summarise and analyse the lecturer hours in TUDublin.")
    print("Two input files are needed here,the first one consists of the raw timetable data.")
    print("The seond input file consists of the contracted hours a lecturer is set to work.")
    print("The file pathway of these files can be entered in the .txt file or here.")
    print("Please ensure that the file is a .xlsx and include .xlsx in the full pathway.")
    print("Also ensure that the first and second file pathway is correct when using 'default'.")
    print("!===================================================================================!\n")

def file_setup(counter):
    """
    file_setup():
    This function is used to define the path of the syllabus plus excel file that the data will be extracted
    from this is the first input file. It is also used to define the file path of the contract hours file, this
    is the second input file. In this function, if there are no previous file paths defined in the .txt file
    then the user will be asked if they want to make the file path they entered when they run the program the
    default file location. There is also functionality in place so if the user wants to enter a new default
    location, they can enter the file path and should be prompted if they want to add this location to the .txt
    file. This way the user does not need to constantly enter the file path of the excel file.The file location
    will be passed to the file_sort() function if this function detects that the path is incorrect the application
    will terminate.
    """
    #loop infinitely until incorrect file is detected or correct file is detected
    while True:
        try:
            #check which file the user is entering based on the parameter value passed in
            if counter == 0:
                file_loc = input(r"Please enter pathway of First excel file here i.e., (C:\Users\JohnDoe\SyllabusPlusfile.xlsx) or type default if pathway has already been set for the syllabus plus output file:")
            if counter == 1:
                file_loc = input(r"Please enter pathway of Second excel file here i.e., (C:\Users\JohnDoe\ContractHoursfile.xlsx) or type default if pathway has already been set for the lecturer contract hours file:")
            print('\n')
            #open the text file that stores the file locations and read lines inside it
            xlfile = open("timetablelocation.txt","r")
            readfile = xlfile.readlines()
            #if .txt file is empty user must can enter and save a file location path 
            if len(readfile) <= 0:
                store_def_loc = input("Would you like to make this the default file location? Yes or No? ")
                if "yes" in store_def_loc.lower():
                    save_loc = open("timetablelocation.txt","w")
                    save_loc.write(file_loc)
                    save_loc.close()
                else:
                    file_loc = file_loc.rstrip()
                    xlfile.close()
            #user can add new file path if a file location is entered and .txt file is more than or equal 1
            elif "\\" in str(file_loc) and len(readfile) >= 1 :
                    addpath = input("Would you like to save this file location?")
                    if 'yes' in addpath.lower():
                        save_loc = open("timetablelocation.txt","w")
                        save_loc.write(file_loc)
                        file_loc = file_loc.rstrip()
                    else:
                        file_loc = file_loc.rstrip()
            #check to see the file paths in .txt file and display, allow user to choose which file path they want
            elif "default" in file_loc.lower() and len(readfile) >= 1:
                if len(readfile) >= 1:
                    counter = 0
                    for i in readfile:
                        print(counter," - ",i)
                        counter += 1
                    while True:
                        try:
                            whichloc = int(input("\nMore than 1 default location detected, please specify which file location to use 0, 1 or 2 etc. : "))
                            if isinstance(whichloc,int):
                                break
                        except ValueError:
                            print("Please ensure that you enter only the numbers beside each of the defined default file locations!")
                        finally:
                            print("Default file path location entered:",whichloc)
                    a = readfile[whichloc].rstrip()
                    file_loc = a
            #check if there is one file path in the .txt if so make it default file path otherwise if 0 let user know and close file
            elif len(readfile) == 1 :
                    file_loc = readfile[0].rstrip()
            elif len(file_loc) == 0:
                print("No default file location found or set")
            xlfile.close()
            #check the file path is distinguished and valid, break infinite loop, else prompt that file is wrong and end program
            if len(file_loc) > 0:
                print("File path accepted!")
                break
            else:
                print("Error occurred retrieving file path, application terminating!")
                sys.exit()
        #catch FileNotFoundErrors and prompt the user instead with a message
        except FileNotFoundError:
            print("Error! Incorrect pathway or pathway not found please try again")
    #when loop is broken return file path back to main
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
    #reads in excel file and places data in appropriate columns and formatted to more usable data
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
    #loops through the split names in dataframe pairs names based on first and last name and added to list
    for i in range(0,len(names)):
        namepairs = list(zip(names[i][::2],names[i][1::2]))
        lst.append(namepairs)
    #dataframe rearranged to accomodate for new name pairs, data is then unstacked and index is reset
    df["Staff Names"] = lst
    df = df.explode('Staff Names').reset_index(drop=True)
    #characters are removed again and columns are reordered, duplicate value instances removed and made a new variable
    df["Staff Names"] = df["Staff Names"].astype(str).replace(r"(?<!\w)'|'(?!\w)|[()]",'',regex=True)
    df["Staff Names"] = df["Staff Names"].astype(str).replace(r'[""]','',regex=True)
    df.index += 1
    df_order = df[["Staff Names","Scheduled Start Time","Duration","Availability","Module Name","Teaching Week Pattern","Number Of Teaching Weeks"]]
    uniqlst = df["Staff Names"].unique()
    #values in unique list are sorted alphabetically and returned to main along with ordered dataframe 
    uniqlst.sort()
    return df_order,uniqlst

def timedelta_to_float(totalseconds):
    """
    timedelta_to_float():
    In this function, the total seconds of a time delta value is found using the total_seconds() module
    available in python. When the time delta value is passed to this function it coverts it to total seconds
    then using defined variables containing the total seconds in a day and total seconds in an hour, the total
    seconds is then first divided by total seconds in a day to find if total seconds is greater than a day
    and then total seconds modulus day total seconds and divided by hours total seconds gives you the total
    remaining hours a lecturer has worked. Adding the converted day and hours value gives the total hours a
    lecturer has completed, this is then rounded to 2 decimal places at the end and passed back to the function
    that called this function.
    """
    #created constant values and used to perform calculations on total seconds of passed in parameter
    day_ttl_seconds = 86400
    hours_ttl_seconds = 3600
    get_ttlseconds = totalseconds.total_seconds()
    days_cnvrt_flt = int(get_ttlseconds/day_ttl_seconds)*24
    hrs_cnvrt_flt = (get_ttlseconds%day_ttl_seconds)/hours_ttl_seconds
    ttlhrs_asfloat = days_cnvrt_flt + hrs_cnvrt_flt
    ttlhrs_2deciml_p = round(ttlhrs_asfloat,2)
    #calculated total hours as a float returned to function 
    return ttlhrs_2deciml_p
    
  
def process_sem1_data(dataframe, uniqlst):
    """
    process_sem1_data():
    In this function, first lists are created and will be used to hold the final processed data of this function.
    Before data is processed characters are removed or changed in the data frame column that is going to be processed,
    this is to make the data easier to parse and read in. There is then an outside for loop this will loop through all
    the unique name values processed earlier in this program starting at index 0 then variables are defined which will
    be used to aid in the data processing, variables are initialised with Timedelta(0) to allow it to hold the data
    type of the dataframe column, if variable values are not initialised as Timedelta(0) then it means that the variable
    is used to store a fixed value. In the inside for loop it then cycles through the lecturer names within the dataframe
    column 'Staff Names', the dataframe column contains many duplicate names but each with a different value. In the for
    loop it checks using a string or regex for certain key words or patterns in the dataframe column all related to semester
    1 including 'semester 1','weeks n-x', 'term n' etc. Then if the data read in passes the condition check there is another
    if to capture data that are less than 13 weeks as calculations need to be applied to convert the hours back to 13 week
    basis. In the calculations below regardless of if they work a 13 week basis or not, basically hours for lecturers name
    equal to the outside for loop at index 'x' are added together and applied and added a night factor of 0.25 to hours if
    it is detected that those hours are more than the 18:00:00 hour threshold. After all hours are added and calculated they
    are then converted to float and added to the appropriate list created at the start of the program. Then these lists are
    passed back to the main function.
    """
    sem1_lst = []
    sem1_unschd_lst = []
    sem1_s1day_lst = []
    sem1_s1night_lst = []
    #formats the dataframe to allow for easier data parsing
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"(-)",',',regex=True)
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"\(",'',regex=True)
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"\)",'',regex=True)
    #use of timedelta reference: https://www.geeksforgeeks.org/python-datetime-timedelta-class/
    for a in range(0, len(uniqlst)):
        night_time = pd.Timedelta(hours=18, minutes=0, seconds=0)
        night_factor = 0.25
        calc_s1day = pd.Timedelta(0)
        unsched_hrs = pd.Timedelta(0)
        s1_day = pd.Timedelta(0)
        s1_night = pd.Timedelta(0)
        control = pd.Timedelta(0)
        counter = pd.Timedelta(0)
        counter1 = pd.Timedelta(0)
        nightcount = pd.Timedelta(0)
        totalhrs = pd.Timedelta(0)
        wks_calc_s1day = pd.Timedelta(0)
        wks_calc_s1night = pd.Timedelta(0)
        wks_convrt_calc_s1night = pd.Timedelta(0)
        wks_unsched_hrs = pd.Timedelta(0)
        wks_s1day = pd.Timedelta(0)
        wks_s1night = pd.Timedelta(0)
        wks_counter = pd.Timedelta(0)
        wks_counter1 = pd.Timedelta(0)
        wks_nightcount = pd.Timedelta(0)
        wks_totalhrs = pd.Timedelta(0)
        wks_realhrs = pd.Timedelta(0)
        for i in range(1, len(dataframe["Staff Names"])):
            #reference: https://medium.com/factory-mind/regex-tutorial-a-simple-cheatsheet-by-examples-649dc1c3f285
            #if statement finds semester, weeks and terms all related to semester 1 using regex search
            if (uniqlst[a] in dataframe["Staff Names"][i] and
                    (("Semester 1" in str(dataframe["Availability"][i]) or "Term 1" in str(dataframe["Availability"][i]))
                     or (str(dataframe["Availability"][i]) == '0' and (re.search(r"\b([4-9]|1[0-6])\b", str(dataframe["Teaching Week Pattern"][i]))))
                     or (re.search(r"Weeks\s+([4-9]|1[0-6])\b", str(dataframe["Availability"][i])))
                     or (re.search(r"Week\s+([4-9]|1[0-6])\b", str(dataframe["Availability"][i]))))):
                #condition statment uses regex to find weeks individual or grouped from range 4 - 16 by matching patterns in the data at index i
                if ((re.search(r"Weeks\s+([4-9]|1[0-6])\b", str(dataframe["Availability"][i])))
                    or (str(dataframe["Availability"][i]) == '0' and (re.search(r"\b([4-9]|1[0-6])\b", str(dataframe["Teaching Week Pattern"][i]))))
                    or "Term 1" in str(dataframe["Availability"][i])
                     or (re.search(r"Week\s+([4-9]|1[0-6])\b", str(dataframe["Availability"][i])))):
                    #variables initialised to hold relevant data concerning weeks
                    nOf_teaching_wks = int(dataframe["Number Of Teaching Weeks"][i])
                    wks_counter += pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    wks_counter1 = pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    wks_start_time = pd.Timestamp(dataframe["Scheduled Start Time"][i])
                    wks_sched_start = pd.Timedelta(hours=wks_start_time.hour, minutes=wks_start_time.minute)
                    wks_sched_end = wks_counter1 + wks_sched_start
                    wks_totalhrs = wks_counter
                    wks_convrt_13 = pd.Timedelta(0)
                    wks_nighthrs = pd.Timedelta(0)
                    #condition statements below check for specific characteristics and adds up hours together respectively
                    if '00:00:00' in str(wks_start_time) and nOf_teaching_wks >= 13:
                        wks_unsched_hrs += wks_counter1
                    if wks_sched_end <= night_time and '00:00:00' not in str(wks_start_time) and nOf_teaching_wks >= 13:
                            wks_s1day += wks_counter1
                    if wks_sched_end > night_time:
                        if wks_sched_start >= night_time:
                            wks_sum1 = wks_sched_end - wks_sched_start
                            wks_nighthrs = wks_sum1 * night_factor
                            wks_calc_s1night = wks_sum1
                            if '00:00:00' not in str(wks_start_time) and nOf_teaching_wks >= 13:
                                wks_s1night += wks_sum1
                        elif wks_sched_start < night_time:
                            wks_calc_s1day = night_time - wks_sched_start
                            wks_sum2 = wks_sched_end - night_time
                            wks_nighthrs = wks_sum2 * night_factor
                            wks_calc_s1night = wks_sum2
                            if '00:00:00' not in str(wks_start_time) and nOf_teaching_wks >= 13:
                                wks_s1night += wks_sum2
                        wks_nightcount += wks_nighthrs
                    #condition statements below calculate the total and individual hours for specific factors and adds values up respectively
                    if wks_nightcount != control:
                        wks_totalhrs = wks_nightcount + wks_counter
                    if wks_calc_s1day != control and nOf_teaching_wks >= 13 and '00:00:00' not in str(wks_start_time):
                        wks_s1day += wks_calc_s1day
                    if nOf_teaching_wks != control and nOf_teaching_wks < 13:
                        wks_convrt_13 = wks_counter1 * (nOf_teaching_wks/13)
                        wks_s1day_cvrt = wks_calc_s1day * (nOf_teaching_wks/13)
                        wks_convrt_calc_s1night = wks_calc_s1night * (nOf_teaching_wks/13) 
                        if wks_s1day != control:
                            wks_s1day += wks_s1day_cvrt
                        if wks_nightcount != control:
                            wks_convrt_13_night = wks_nighthrs * (nOf_teaching_wks/13)
                            wks_convrt_13 = wks_convrt_13 + wks_convrt_13_night
                        wks_realhrs += wks_convrt_13
                        if '00:00:00' in str(wks_start_time):
                            wks_unsched_hrs += wks_convrt_13
                        if wks_sched_end <= night_time and '00:00:00' not in str(wks_start_time):
                            wks_s1day += wks_convrt_13
                        if (wks_sched_end > night_time and '00:00:00' not in str(wks_start_time) and wks_convrt_calc_s1night != control):
                            wks_s1night += wks_convrt_calc_s1night
                #else condition catches remaining data that does not pass condition statement
                else:
                    #process below follows the same calculation as above condition statement but has less parameters to consider
                    counter += pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    counter1 = pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    nOf_teaching_wks = int(dataframe["Number Of Teaching Weeks"][i])
                    start_time = pd.Timestamp(dataframe["Scheduled Start Time"][i])
                    sched_start = pd.Timedelta(hours=start_time.hour, minutes=start_time.minute)
                    sched_end = counter1 + sched_start
                    totalhrs = counter
                    nighthrs = pd.Timedelta(0)
                    if '00:00:00' in str(start_time) and nOf_teaching_wks >= 13:
                        unsched_hrs += counter1
                    if sched_end <= night_time and '00:00:00' not in str(start_time) and nOf_teaching_wks >= 13:
                        s1_day += counter1 
                    if sched_end > night_time:
                        if sched_start >= night_time:
                            sum1 = sched_end - sched_start
                            nighthrs = sum1 * night_factor
                            if '00:00:00' not in str(start_time):
                                s1_night += sum1   
                        elif sched_start < night_time:
                            calc_s1day = night_time - sched_start
                            sum2 = sched_end - night_time
                            nighthrs = sum2 * night_factor
                            if '00:00:00' not in str(start_time):
                                s1_night += sum2
                        nightcount += nighthrs
            if nightcount != control:
                totalhrs = nightcount + counter
        #final checks are completed to add week equivalent total hours and total hours
        if wks_totalhrs != control:
            totalhrs = totalhrs + wks_realhrs
        if wks_s1night  != control:
            s1_night = s1_night + wks_s1night
        if wks_s1day != pd.Timedelta(0) or calc_s1day != pd.Timedelta(0):
            s1_day = s1_day + wks_s1day + calc_s1day 
        if wks_unsched_hrs != pd.Timedelta(0):
            unsched_hrs = unsched_hrs + wks_unsched_hrs
        #values converted to float added to list and returned back to main
        s1day_ttlhrs_2dp = timedelta_to_float(s1_day)
        s1night_ttlhrs_2dp = timedelta_to_float(s1_night)
        unschd_ttlhrs_2dp = timedelta_to_float(unsched_hrs)
        ttlhrs_2dp = timedelta_to_float(totalhrs)
        sem1_s1day_lst.append(s1day_ttlhrs_2dp)
        sem1_s1night_lst.append(s1night_ttlhrs_2dp)
        sem1_unschd_lst.append(unschd_ttlhrs_2dp)
        sem1_lst.append(ttlhrs_2dp)
    return sem1_lst,sem1_unschd_lst,sem1_s1day_lst,sem1_s1night_lst

def process_sem2_data(dataframe, uniqlst):
    """
    process_sem2_data():
    In this function, lists are created and will be used to hold the final processed data of this function. Before data is
    processed characters are removed or changed in the data frame column that is going to be processed, this is to make the
    data easier to parse and read in. Just like the process_sem2_data() function There is then an outside for loop this will
    loop through all the unique name values processed earlier in this program starting at index 0 then variables are defined
    which will be used to aid in the data processing, variables are initialised with Timedelta(0) to allow it to hold the data
    type of the dataframe column, if variable values are not initialised as Timedelta(0) then it means that the variable is used
    to store a fixed value. In the inner for loop it then cycles through the lecturer names within the dataframe column 'Staff Names'.
    The dataframe column contains many duplicate names but each with a different value. Similarly to the function to process
    semester 1 data the only difference in this function is that there are more weeks and parameters to account for as term 2
    and term 3 also fall into weeks of semester 2 and are counted as semester 2 in accordance to the data in the output report.
    In the for loop it checks using a string or regex for certain key words or patterns in the dataframe column all related to
    semester 1 including 'semester 2','weeks n-x', 'term n' etc. Then if the data read in passes the condition check there is
    another if to capture data that are less than 13 weeks as calculations need to be applied to convert the hours back to 13
    week basis. In the calculations below regardless of if they work a 13 week basis or not, basically hours for lecturers name
    equal to the outside for loop at index 'x' are added together and applied and added a night factor of 0.25 to hours if
    it is detected that those hours are more than the 18:00:00 hour threshold. After all hours are added and calculated they
    are then converted to float and added to the appropriate list created at the start of the program. Then these lists are
    passed back to the main function.
    """
    sem2_lst = []
    sem2_unschd_lst = []
    sem2_s2day_lst = []
    sem2_s2night_lst = []
    #formats the dataframe to allow for easier data parsing
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"(-)",',',regex=True)
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"\(",'',regex=True)
    dataframe["Availability"] = dataframe["Availability"].astype(str).replace(r"\)",'',regex=True)
    #use of timedelta reference: https://www.geeksforgeeks.org/python-datetime-timedelta-class/
    for a in range(0, len(uniqlst)):
        night_time = pd.Timedelta(hours=18, minutes=0, seconds=0)
        night_factor = 0.25
        calc_s2day = pd.Timedelta(0)
        unsched_hrs = pd.Timedelta(0)
        s2_day = pd.Timedelta(0)
        s2_night = pd.Timedelta(0)
        control = pd.Timedelta(0)
        counter = pd.Timedelta(0)
        counter1 = pd.Timedelta(0)
        nightcount = pd.Timedelta(0)
        totalhrs = pd.Timedelta(0)
        wks_calc_s2day = pd.Timedelta(0)
        wks_calc_s2night = pd.Timedelta(0)
        wks_convrt_calc_s2night = pd.Timedelta(0)
        wks_unsched_hrs = pd.Timedelta(0)
        wks_s2day = pd.Timedelta(0)
        wks_s2night = pd.Timedelta(0)
        wks_counter = pd.Timedelta(0)
        wks_counter1 = pd.Timedelta(0)
        wks_nightcount = pd.Timedelta(0)
        wks_totalhrs = pd.Timedelta(0)
        wks_realhrs = pd.Timedelta(0)
        for i in range(1, len(dataframe["Staff Names"])):
            #\b is a word boundary which essentially allows only position between the boundary defined to be matched meaning if anything follows the number it is not matched.
            #reference: https://medium.com/factory-mind/regex-tutorial-a-simple-cheatsheet-by-examples-649dc1c3f285
            #if statement finds semester, weeks and terms all related to semester 2 using regex search
            if (uniqlst[a] in dataframe["Staff Names"][i] and
                    (("Semester 2" in str(dataframe["Availability"][i]) or "Term 2" in str(dataframe["Availability"][i]))
                     or "Semester 1&2" in str(dataframe["Availability"][i]).lstrip()
                     or "Term 3" in str(dataframe["Availability"][i])
                     or (str(dataframe["Availability"][i]) == '0' and (re.search(r"(1[8-9]|2[0-9]|3[0-9]|4[0-5])", str(dataframe["Teaching Week Pattern"][i]))))
                     or (re.search(r"(1[8-9]|2[0-9]|3[0-9]4[0-5])", str(dataframe["Availability"][i])))
                     or (re.search(r"Weeks\s+(1[8-9]|2[0-9]|3[0-9]|4[0-5])\b", str(dataframe["Availability"][i])))
                     or (re.search(r"Week\s+(1[8-9]|2[0-9]|3[0-9]|4[0-5])\b", str(dataframe["Availability"][i]))))):
                #section uses regex to find weeks individual or grouped from range 18 - 45 by matching patterns in the data at index i
                if ((re.search(r"Weeks\s+(1[8-9]|2[0-9]|3[0-9]|4[0-5])\b", str(dataframe["Availability"][i])))
                    or (str(dataframe["Availability"][i]) == '0' and (re.search(r"(1[8-9]|2[0-9]|3[0-9]|4[0-5])", str(dataframe["Teaching Week Pattern"][i]))))
                    or "Term 2" in str(dataframe["Availability"][i])
                     or "Term 3" in str(dataframe["Availability"][i])
                     or (re.search(r"Week\s+(1[8-9]|2[0-9]|3[0-9]|4[0-5])\b", str(dataframe["Availability"][i])))):
                    #variables initialised to hold relevant data concerning weeks
                    nOf_teaching_wks = int(dataframe["Number Of Teaching Weeks"][i])
                    wks_counter += pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    wks_counter1 = pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    wks_start_time = pd.Timestamp(dataframe["Scheduled Start Time"][i])
                    wks_sched_start = pd.Timedelta(hours=wks_start_time.hour, minutes=wks_start_time.minute)
                    wks_sched_end = wks_counter1 + wks_sched_start
                    wks_totalhrs = wks_counter
                    wks_convrt_13 = pd.Timedelta(0)
                    wks_nighthrs = pd.Timedelta(0)
                    #condition statements below check for specific characteristics and adds up hours together respectively
                    if '00:00:00' in str(wks_start_time) and nOf_teaching_wks >= 13:
                        wks_unsched_hrs += wks_counter1
                    if wks_sched_end <= night_time and '00:00:00' not in str(wks_start_time) and nOf_teaching_wks >= 13:
                        wks_s2day += wks_counter1
                    if wks_sched_end > night_time:
                        if wks_sched_start >= night_time:
                            wks_sum1 = wks_sched_end - wks_sched_start
                            wks_nighthrs = wks_sum1 * night_factor
                            wks_calc_s2night = wks_sum1
                            if '00:00:00' not in str(wks_start_time) and nOf_teaching_wks >= 13:
                                wks_s2night += wks_sum1
                        elif wks_sched_start < night_time:
                            wks_calc_s2day = night_time - wks_sched_start
                            wks_sum2 = wks_sched_end - night_time
                            wks_nighthrs = wks_sum2 * night_factor
                            wks_calc_s1night = wks_sum2
                            if '00:00:00' not in str(wks_start_time) and nOf_teaching_wks >= 13:
                                wks_s2night += wks_sum2
                        wks_nightcount += wks_nighthrs
                    #condition statements below calculate the total and individual hours for specific factors and adds values up respectively
                    if wks_nightcount != control:
                        wks_totalhrs = wks_nightcount + wks_counter
                    if wks_calc_s2day != pd.Timedelta(0) and nOf_teaching_wks >= 13 and '00:00:00' not in str(wks_start_time):
                        wks_s2day += wks_calc_s2day                       
                    if nOf_teaching_wks != control and nOf_teaching_wks < 13:
                        wks_convrt_13 = wks_counter1 * (nOf_teaching_wks/13)
                        wks_s2day_cvrt = wks_calc_s2day * (nOf_teaching_wks/13)
                        wks_convrt_calc_s2night = wks_calc_s2night * (nOf_teaching_wks/13)
                        if wks_s2day != control:
                            wks_s2day += wks_s2day_cvrt
                        if wks_nightcount != control:
                            wks_convrt_13_night = wks_nighthrs * (nOf_teaching_wks/13)
                            wks_convrt_13 = wks_convrt_13 + wks_convrt_13_night
                        wks_realhrs += wks_convrt_13
                        if '00:00:00' in str(wks_start_time):
                            wks_unsched_hrs += wks_convrt_13
                        if wks_sched_end <= night_time and '00:00:00' not in str(wks_start_time):
                            wks_s2day += wks_convrt_13
                        if (wks_sched_end > night_time and '00:00:00' not in str(wks_start_time) and wks_convrt_calc_s2night != control):
                            wks_s2night += wks_convrt_calc_s2night
                            
                #else condition catches remaining data that does not pass condition statement
                else:
                    #process below follows the same calculation as above condition statement but has less parameters to consider
                    counter += pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    counter1 = pd.Timedelta(hours=dataframe["Duration"][i].hour,minutes=dataframe["Duration"][i].minute)
                    start_time = pd.Timestamp(dataframe["Scheduled Start Time"][i])
                    nOf_teaching_wks = int(dataframe["Number Of Teaching Weeks"][i])
                    sched_start = pd.Timedelta(hours=start_time.hour, minutes=start_time.minute)
                    sched_end = counter1 + sched_start
                    totalhrs = counter
                    nighthrs = pd.Timedelta(0)
                    if '00:00:00' in str(start_time) and nOf_teaching_wks >= 13:
                        unsched_hrs += counter1
                    if sched_end <= night_time and '00:00:00' not in str(start_time) and nOf_teaching_wks >= 13:
                        s2_day += counter1 
                    if sched_end > night_time:
                        if sched_start >= night_time:
                            sum1 = sched_end - sched_start
                            nighthrs = sum1 * night_factor
                            if '00:00:00' not in str(start_time):
                                s2_night += sum1 
                        elif sched_start < night_time:
                            calc_s2day = night_time - sched_start
                            sum2 = sched_end - night_time
                            nighthrs = sum2 * night_factor
                            if '00:00:00' not in str(start_time):
                                s2_night += sum2
                        nightcount += nighthrs
            if nightcount != control:
                totalhrs = nightcount + counter
        #final checks are completed to add week equivalent total hours and total hours
        if wks_totalhrs != control:
            totalhrs = totalhrs + wks_realhrs
        if wks_s2night != control:
            s2_night = s2_night + wks_s2night
        if wks_s2day != pd.Timedelta(0) or calc_s2day != pd.Timedelta(0):
            s2_day = s2_day + wks_s2day + calc_s2day 
        if wks_unsched_hrs != pd.Timedelta(0):
            unsched_hrs = unsched_hrs + wks_unsched_hrs
        #values converted to float added to list and returned back to main
        s2day_ttlhrs_2dp = timedelta_to_float(s2_day)
        s2night_ttlhrs_2dp = timedelta_to_float(s2_night)
        unschd_ttlhrs_2dp = timedelta_to_float(unsched_hrs)
        ttlhrs_2dp = timedelta_to_float(totalhrs)
        sem2_s2day_lst.append(s2day_ttlhrs_2dp)
        sem2_s2night_lst.append(s2night_ttlhrs_2dp)
        sem2_unschd_lst.append(unschd_ttlhrs_2dp)
        sem2_lst.append(ttlhrs_2dp)
    return sem2_lst,sem2_unschd_lst,sem2_s2day_lst,sem2_s2night_lst

def data_analysis(file_location,sem1_lst,sem2_lst,sem1_unschd_lst,sem2_unschd_lst,sem1_s1day_lst,sem2_s2day_lst,sem1_s1night_lst,sem2_s2night_lst,uniqlst):
    """
    data_analysis():
    In this function the location of the second input file is read in, which holds the contract hours for lecturers that work
    in TUDublin. The data from this file is placed inside a dataframe and columns labelled accordingly. Then using the processed
    data for semester 1 and 2 another dataframe is created, and each list item is placed into an appropriately named dataframe
    column and reordered. Spaces are removed from both dataframes to allow the program to compare the names correctly within each
    dataframe, lists are also created which will be used to create the final dataframe of the completed analysed data. A for loop
    then cycles through both columns labelled 'Lecturers1' and 'Lecturers' only lecturers relevant to TUDublin are filtered out.
    Then total hours,night and day hours for lecturers in TUDublin are also filtered out and placed into the correct list, another
    for loop is implemented to apply the calculations for the under or over hours calculation, as well as multiplying the total hours
    of a lecturer by 13 to get the total hours over the year. This is then also added to the appropriate list, the final for loop is
    used to get the total under or over hours and added to the appropriate list. Then a new dataframe is created to store the new data,
    then two different variants are created of this final dataframe one includes the total hours and the other does not. The two variants
    are then passed back to the main function.
    """
    #two dataframes created from the second input file and the processed data of semester 1 and 2
    xl = pd.ExcelFile(file_location)
    read1 = pd.read_excel(xl, 'Lecturers',usecols="A,B,C", names=["Lecturers1","S1 Hours","S2 Hours"])
    contracth_df = pd.DataFrame(read1)
    realtime_df = pd.DataFrame(sem1_lst,columns=["S1 Hours"])
    realtime_df['S1 Unsch'] = sem1_unschd_lst
    realtime_df['S2 Unsch'] = sem2_unschd_lst
    realtime_df['S1 Day'] = sem1_s1day_lst
    realtime_df['S1 Night'] = sem1_s1night_lst
    realtime_df['S2 Day'] = sem2_s2day_lst
    realtime_df['S2 Night'] = sem2_s2night_lst
    realtime_df['S2 Hours'] = sem2_lst
    realtime_df['Lecturers'] = uniqlst
    realtime_df = realtime_df[['Lecturers','S1 Hours','S2 Hours','S1 Unsch','S1 Day','S1 Night','S2 Unsch','S2 Day','S2 Night']]
    realtime_df['Lecturers'] = realtime_df['Lecturers'].astype(str).replace(r"(  )",'',regex=True)
    contracth_df['Lecturers1'] = contracth_df['Lecturers1'].astype(str).replace(r"( )",'',regex=True)
    #lists created are used to hold data that will be used to create the final dataframe
    staff_hrstotalhrsS1 = []
    staff_hrstotalhrsS2 = []
    staff_unschdhrsS1 = []
    staff_unschdhrsS2 = []
    staff_S1Day = []
    staff_S2Day = []
    staff_S1Night = []
    staff_S2Night = []
    staff_hrsS1_undr_over = []
    staff_hrsS2_undr_over = []
    total_under_over = []
    yearhrs = []
    print("\n")
    #loops through and compares names in unique set and second input file names
    for i in range(0,len(contracth_df['Lecturers1'])):
        #if names match extract the associated processed data and add to list appropriate
        if contracth_df['Lecturers1'][i] in str(realtime_df['Lecturers']).lstrip():
            indx = realtime_df.index[realtime_df['Lecturers'] == contracth_df['Lecturers1'][i]][0]
            staff_hrstotalhrsS1.append(realtime_df['S1 Hours'][indx])
            staff_hrstotalhrsS2.append(realtime_df['S2 Hours'][indx])
            staff_unschdhrsS1.append(realtime_df['S1 Unsch'][indx])
            staff_unschdhrsS2.append(realtime_df['S2 Unsch'][indx])
            staff_S1Day.append(realtime_df['S1 Day'][indx])
            staff_S2Day.append(realtime_df['S2 Day'][indx])
            staff_S1Night.append(realtime_df['S1 Night'][indx])
            staff_S2Night.append(realtime_df['S2 Night'][indx])
        #if names dont match add a 0 to that index of list appropriate
        else:
            staff_hrstotalhrsS1.append(0)
            staff_hrstotalhrsS2.append(0)
            staff_unschdhrsS1.append(0)
            staff_unschdhrsS2.append(0)
            staff_S1Day.append(0)
            staff_S2Day.append(0)
            staff_S1Night.append(0)
            staff_S2Night.append(0)
    #loops through names in dataframe and performs calculations to get wanted value and add to list appropriate  
    for i in range(0,len(contracth_df['Lecturers1'])):
        s1_chs_sub_ttlhrs = float(staff_hrstotalhrsS1[i]) - float(contracth_df['S1 Hours'][i])
        s2_chs_sub_ttlhrs = float(staff_hrstotalhrsS2[i]) - float(contracth_df['S2 Hours'][i])
        yeartotal = (float(staff_hrstotalhrsS1[i])*13)+(float(staff_hrstotalhrsS2[i])*13)
        staff_hrsS1_undr_over.append(s1_chs_sub_ttlhrs)
        staff_hrsS2_undr_over.append(s2_chs_sub_ttlhrs)
        yearhrs.append(yeartotal)
    #loop through names in dataframe and perform calculation to get wanted sum value and add to list appropriate
    for i in range(0,len(contracth_df['Lecturers1'])):
        totalundrover = float(staff_hrsS1_undr_over[i]) + float(staff_hrsS2_undr_over[i])
        total_under_over.append(totalundrover)

    #create final dataframe and append correct lists to named columns
    customer_outputRep = pd.DataFrame()
    customer_outputRep['Lecturers'] = contracth_df['Lecturers1']
    customer_outputRep['CHS1'] = contracth_df['S1 Hours']
    customer_outputRep['S1 Total Hours'] = staff_hrstotalhrsS1
    customer_outputRep['S1 Over'] = staff_hrsS1_undr_over
    customer_outputRep['CHS2'] = contracth_df['S2 Hours']
    customer_outputRep['S2 Total Hours'] = staff_hrstotalhrsS2
    customer_outputRep['S2 Over'] = staff_hrsS2_undr_over
    customer_outputRep['Over Hrs'] = total_under_over
    customer_outputRep['Year'] = yearhrs
    customer_outputRep['S1 Unsch'] = staff_unschdhrsS1
    customer_outputRep['S2 Unsch'] = staff_unschdhrsS2
    customer_outputRep['S1 Day'] = staff_S1Day
    customer_outputRep['S2 Day'] = staff_S2Day
    customer_outputRep['S1 Night'] = staff_S1Night
    customer_outputRep['S2 Night'] = staff_S2Night
    #two versions of dataframe created one holds additional columns and other does not which are both returned to main
    customer_outputRepA = customer_outputRep[['Lecturers','CHS1','S1 Over','CHS2','S2 Over','Over Hrs','Year','S1 Unsch','S1 Day','S1 Night','S2 Unsch','S2 Day','S2 Night']]
    customer_outputRepB = customer_outputRep[['Lecturers','CHS1','S1 Total Hours','S1 Over','CHS2','S2 Total Hours','S2 Over','Over Hrs','Year','S1 Unsch','S1 Day','S1 Night','S2 Unsch','S2 Day','S2 Night']]
    print(customer_outputRepA)
    return customer_outputRepA,customer_outputRepB

def exportToExcel(customer_outputRepA,customer_outputRepB):
    """
    exportToExcel():
    In this function the two dataframes that are passed to it are wrote to a new excel file, two new datasheets are created
    one for each of the dataframe that is wrote to it. 
    """
    #using pandas module ExcelWriter() write to new excel file and add two dataframes on seperate datasheets
    with pd.ExcelWriter('Timetable_Analysis_Report.xlsx') as writer:
        customer_outputRepA.to_excel(writer, sheet_name='Summary')
        customer_outputRepB.to_excel(writer, sheet_name='Summary Total Hours Included')
    

def main():
    """
    main():
    In this function the configurations are set for the display of the dataframe on the IDE for troubleshooting purposes. The counter
    here is passed to the file_setup() function this allows the file_setup() function to distinguish between the first and the second
    file path, if there is nothing set in default.There is an if statement in place to check if the file path shows a correct excel file
    to parse. Then the file_sort() function is called to modify the data to a more usable format, after it is created into a more usable
    format. It is passed to the process_semN_data() function, this returns the data that will be compared to the second input file that
    holds the contract hours for staff explicitly in TUDublin. Counter is then set to 1 and sent to file_setup() function again to allow
    it to distinguish that this is the second input file. Then the data_analysis() function is called passed the appropriate parameters
    to allow it to return the dataframes holding the information requested by the customer. Then the exportToExcel() function is called
    to create a excel file that displays the customers requested data.
    """
    set_configs()
    counter = 0
    file_path1 = file_setup(counter)
    dataframe, uniqlst = file_sort(file_path1)
    if len(dataframe) > 0 or len(uniqlst) > 0:
        sem1_lst,sem1_unschd_lst,sem1_s1day_lst,sem1_s1night_lst = process_sem1_data(dataframe, uniqlst)
        print("\n")
        sem2_lst,sem2_unschd_lst,sem2_s2day_lst,sem2_s2night_lst = process_sem2_data(dataframe, uniqlst)
    else:
        print("Error occurred retrieving data application terminating")
        sys.exit()
    counter = 1
    file_path2 = file_setup(counter)
    customer_outputRepA,customer_outputRepB = data_analysis(file_path2,sem1_lst,sem2_lst,sem1_unschd_lst,sem2_unschd_lst,sem1_s1day_lst,sem2_s2day_lst,sem1_s1night_lst,sem2_s2night_lst,uniqlst)
    exportToExcel(customer_outputRepA,customer_outputRepB)
if __name__== "__main__":
    main()
    
