import pandas as pd
import streamlit as st
import numpy as np
import openpyxl
import datetime
import os
#import pickle
#from datetime import datetime
from datetime import timedelta
## Required for writing to goolge sheet


data_file = ".\Faculty_Planner.xlsx"
data= openpyxl.load_workbook(data_file)

#df = pd.read_excel(data_file,sheet_name = "Schedule")

#############################################

## Faculty sheet Load
SHEET ="Faculty"
df_Faculty=pd.read_excel(data_file,sheet_name = SHEET)
#st.write(df_Faculty)
#############################################

## Schedule sheet Load

SHEET = "Schedule"
df_Schedule=pd.read_excel(data_file,sheet_name = SHEET)
#st.write(df_Schedule)

## Converting "Start Date" and "end Date" columns to DateTime datatype 

df_Schedule['Start Date'] = pd.to_datetime(df_Schedule['Start Date'])
df_Schedule['End Date'] = pd.to_datetime(df_Schedule['End Date'])
#st.write(df_Schedule)
##############################################

## Calendar sheet Load

SHEET = "Calendar"
df_Calendar=pd.read_excel(data_file,sheet_name = SHEET)

## Converting Date column to DateTime datatype 

df_Calendar.Date = pd.to_datetime(df_Calendar.Date)
#st.write(df_Calendar)
#####################################################################

## Function to convert time to numeric type

def time_convert(times):
    cnt=0
    comp = times.split(':')
    cnt+=int(comp[0])
    if comp[1]=='00':
        cnt+=0
    elif comp[1]=='30':
        cnt+=0.5
    if times[-2:] == 'PM':
        cnt+=12
    return cnt  

# Convert "Sart time" and "End time" functions to numeric types for calendar and schedule sheet



#st.write(df_Calendar)
### Reading Ratings ########### 
######## Ratings fetched from Metabase--> Copy paste mannually

SHEET = "Rating"

df_Rating=pd.read_excel(data_file,sheet_name = SHEET)
#### Reading batch data ###########
SHEET = "Batch"
df_Batch=pd.read_excel(data_file,sheet_name = SHEET)
#df_Batch['Latest_Scheduled_Date'] = pd.to_datetime(df_Batch['Latest_Scheduled_Date']

####### Reading Module sequence #########
SHEET = "Modules"
df_Modules=pd.read_excel(data_file,sheet_name = SHEET)

## Reading Weight matrix ############
SHEET = "Weight"
df_Weight=pd.read_excel(data_file,sheet_name = SHEET)
##########################################################


SHEET = 'Schedule_online'

df_Online_Schedule=pd.read_excel(data_file,sheet_name = SHEET)
df_Online_Schedule['Start Date'] = pd.to_datetime(df_Online_Schedule['Start Date'])
df_Online_Schedule['End Time'] = df_Online_Schedule['End Time'].apply(time_convert)
df_Online_Schedule['Start Time'] = df_Online_Schedule['Start Time'].apply(time_convert)
#############################################################
############# Reading online bathes ############
SHEET ='Online_batch'

df_Onlinebatch=pd.read_excel(data_file,sheet_name = SHEET)
#print(df_Onlinebatch)
df_Onlinebatch['Latest_Scheduled_Week'] = pd.to_datetime(df_Onlinebatch['Latest_Scheduled_Week'])
########### Function to write to Google sheet #########


SHEET = 'Schedule_online'

df_Online_Schedule=pd.read_excel(data_file,sheet_name = SHEET)
df_Online_Schedule['Start Date'] = pd.to_datetime(df_Online_Schedule['Start Date'])
df_Online_Schedule['End Time'] = df_Online_Schedule['End Time'].apply(time_convert)
df_Online_Schedule['Start Time'] = df_Online_Schedule['Start Time'].apply(time_convert)
#############################################################
############# Reading online bathes ############
SHEET = 'Online_batch'

df_Onlinebatch=pd.read_excel(data_file,sheet_name = SHEET)
#print(df_Onlinebatch)
df_Onlinebatch['Latest_Scheduled_Week'] = pd.to_datetime(df_Onlinebatch['Latest_Scheduled_Week'])
########### Function to write to Google sheet #########


###########################################################

########### Function to write to Google sheet

###

def write_to_sheet(SHEET,value):
    sheet_write = data.get_sheet_by_name(SHEET)
    sheet_write.append(value) 
    data.save(data_file)

###############################################
################## Function to update the sheet ########################

def update_sheet(place,val,sheet_id):
    sheet_to_update = data.get_sheet_by_name(sheet_id)
    c = place[0]
    col_dict = {'A':1,'B':2,'C':3,'D':4,'E':5,'F':6,'G':7,
    'H':8,'I':9 ,'J':10,'K':11,'L':12,'M':13,'N':14,'O':15,'P':16,
    'Q':17,'R':18,'S':19,'T':20,'U':21,'V':22,'W':23,'X':24,
    'Y':25,'Z':26}
    r = int(place[1:])
    c = col_dict[c]
    for i in range(len(val)):
        sheet_to_update.cell(row=r, column=c).value = val[i]
        c=c+1
    #sheet_to_update[place] = value
    data.save(data_file)

##################################################################


################# Logic Functions #######################
## Function to fetch active batches for the location 
def all_active(loc):
    loc_batches = df_Batch[df_Batch['Location']==loc]
    active_batches = loc_batches[loc_batches['Is_active']=='YES']
    return active_batches.loc[:,['Batch']]
###########################################################

## Function to fetch all online active batches #######
def all_onlineactive():
    active_batches = df_Onlinebatch[df_Onlinebatch['Is_active']=='YES']
    #print('*********************',active_batches.loc[:,['Batch']])
    return active_batches.loc[:,['Batch']]

##################################################################

## Function to fetch the schedule status for the given batch
def next_module_date(batch,location):
    latest_module = df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Latest_scheduled_Module'].iloc[0]
    latest_module_index = df_Modules[df_Modules['Module Name']==latest_module]['Sequence'].iloc[0]
    #next_module_index = str(int(latest_module_index)+1)
    #last_module_index = df_Modules.tail(1)['Sequence'].values[0]
    latest_date = df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Latest_Scheduled_Date'].iloc[0]
    #date_list = pd.date_range(start=latest_date, periods=8, freq = 'D')
    #if next_module_index != last_module_index:
        #next_module = df_Modules[df_Modules['Sequence']==next_module_index]['Module Name'].values[0]
        #next_date = date_list[-1] 
        #return(latest_module,latest_date,next_module,next_date)
    return(latest_module,latest_date)
############################################################################################
##### Function for next module date for online #############
def next_module_date_online(batch):
    latest_module = df_Onlinebatch[df_Onlinebatch['Batch']==batch]['Latest_scheduled_Module'].values[0]
    latest_module_index = df_Modules[df_Modules['Module Name']==latest_module]['Sequence'].values[0]
    #next_module_index = str(int(latest_module_index)+1)
    #last_module_index = df_Modules.tail(1)['Sequence'].values[0]
    latest_date = df_Batch[df_Onlinebatch['Batch']==batch]['Latest_Scheduled_Date'].iloc[0]
    return(latest_module,latest_date)
    #####################################################################
### The function to fetch searching the available faculty for given module at the location
def search_faculty(module_name,location):
    df1 = df_Faculty[df_Faculty[module_name]=='Yes']
    list_faculty = list(df1['Faculty Name'])
    #print(list_faculty)
    dict_faculty={}
    list_fac =[]
    list_loc =[]
    list_wei =[]
    list_int =[]
    for f in list_faculty:
        dict_faculty.update({f:0})
        list_fac.append(f)
        #print("f",f,df1[df1['Faculty Name']==f]['Internal'])
        inter = df1[df1['Faculty Name']==f]['Internal'].iloc[0]
        loc = df1[df1['Faculty Name']==f]['Location'].iloc[0]
        list_loc.append(loc)
        list_int.append(inter)
        if inter.lower() == 'yes':
            w = int(df_Weight[df_Weight['Criteria']=='Internal']['Weight'].iloc[0])
            w1 = dict_faculty[f] 
            dict_faculty[f] = w+w1
        else:
            w = int(df_Weight[df_Weight['Criteria']=='External']['Weight'].iloc[0])
            w1 = dict_faculty[f] 
            dict_faculty[f] = w+w1           
        if loc.lower() == location.lower():
            w = int(df_Weight[df_Weight['Criteria']=='Location']['Weight'].iloc[0])
            w1 = dict_faculty[f] 
            dict_faculty[f] = w+w1
        #print(dict_faculty,inter,sep="|")
    
    for x in dict_faculty.values():
        list_wei.append(x)
    d=dict(Name=list_fac,loc=list_loc,inernal=list_int,weight=list_wei)
    #print(d)
    df_try = pd.DataFrame.from_dict(d)
    #print(df_try)
    return df_try
############################################################################
### Function to check availability of the Faculty

def check_availability(faculty,start_date, end_date, start_time, end_time,resi_type):
    dates=pd.date_range(start= start_date, end=end_date)
    st = time_convert(start_time)
    et = time_convert(end_time)
  
    if resi_type == 'FT':

    #print(dates)
        df1 = df_Calendar[df_Calendar['Faculty']==faculty]
        if df1.shape[0]==0:
            return True
        for d in dates:
            df1 = df1[df1['Date']==d]
            #print(df1)
            if df1.shape[0] != 0:
                df1 = df1[(df1['Start_Time']< et) |  (df1['Start_Time']> st)]
                if df1.shape[0] !=0:
                    return False
                else:
                    flag = True
            else:
                flag = True
        else:
            if flag == True:
                return True
###########################################################################
#######################Function to read schedule of specific batch ###########
def read_schedule(location,batch,sdate=0,ldate=0):
    #print(type(sdate))

    #print(type(ldate))
    df_batch_schedule = df_Schedule[(df_Schedule['Location']==location) & (df_Schedule['Batch']==batch)]
    if sdate!=0 and ldate!=0:
        sdate = pd.to_datetime(sdate)
        ldate = pd.to_datetime(ldate)
        df_batch_schedule = df_batch_schedule[(df_batch_schedule['Start Date'].dt.date>=sdate) & (df_batch_schedule['Start Date'].dt.date<=ldate)]
    return df_batch_schedule
################################################################################

#######################Function to read schedule of specific batch ###########
def read_online_schedule(batch,sdate=0,ldate=0):
    #print(type(sdate))

    #print(type(ldate))
    df_online_batch_schedule = df_Online_Schedule[df_Online_Schedule['Batch']==batch]
    if sdate!=0 and ldate !=0:
        sdate = pd.to_datetime(sdate)
        ldate = pd.to_datetime(ldate)
        df_online_batch_schedule = df_online_batch_schedule[(df_online_batch_schedule['Start Date'].dt.date>=sdate) & (df_online_batch_schedule['Start Date'].dt.date<=ldate)]
    return df_online_batch_schedule
################################################################################
################## Function to read the schedule for faculty #############
def read_faculty_schedule(faculty,sdate,ldate):
        #print(type(sdate))
    sdate = pd.to_datetime(sdate)
    ldate = pd.to_datetime(ldate)
    #print(type(ldate))
    print(df_Schedule)
    df_faculty_schedule = df_Schedule[df_Schedule['Faculty']==faculty]
    print(df_faculty_schedule)
    print(df_faculty_schedule.dtypes)
    df_faculty_schedule = df_faculty_schedule[(df_faculty_schedule['Start Date'].dt.date>=sdate) & (df_faculty_schedule['Start Date'].dt.date<=ldate)]
    return df_faculty_schedule
################## Function to update the sheet ########################

##################################################################
########### Function to update the batch detailes ######################
def set_batch(batch,location,latest_date,latest_module):
    index = df_Batch.index
    condition = (df_Batch['Batch']==batch) & (df_Batch['Location']==location)
    number = index[condition]
    #print(number.tolist())
    #condition =number.tolist()[0]
    x=df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)].index
    s_date = df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Start_date'].iloc[0]
    owner = df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Batch_Owner'].iloc[0]
    if latest_module == 'Case Study':
        active = 'NO'
    else:
        active = 'YES'
    #week = int(df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Week_of_Year'][0])+1
    #print(week)
    number = index[condition]
    row = number.tolist()[0]+2
    #print(row)
    ranges = 'A'+str(row)
    #print(ranges)
    update_sheet(ranges,list([batch, location, str(s_date),active, owner,latest_module,str(latest_date)]),'Batch')
###################################################################################################
######### Function to set online batch data ############
def set_onlinebatch(batch,latest_date,latest_module):
    index = df_Onlinebatch.index
    condition = (df_Onlinebatch['Batch']==batch)
    number = index[condition]
    #print(number.tolist())
    #condition =number.tolist()[0]
    x=df_Onlinebatch[df_Onlinebatch['Batch']==batch]['Start_date'].index
    s_date = df_Onlinebatch[df_Onlinebatch['Batch']==batch]['Start_date'].iloc[0]
    if latest_module == 'Case Study':
        active = 'NO'
    else:
        active = 'YES'
    #week = int(df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Week_of_Year'][0])+1
    #print(week)
    number = index[condition]
    row = number.tolist()[0]+2
    #print(row)
    ranges = 'A'+str(row)
    #print(ranges)
    update_sheet(ranges,list([batch, str(s_date),active, latest_module,str(latest_date)]),'Online_batch')
################################################################

###################### Function to set the batch inactive ####################
def set_inactive(batch,location):
    index = df_Batch.index
    condition = (df_Batch['Batch']==batch) & (df_Batch['Location']==location)
    number = index[condition]
    #print(number.tolist())
    #condition =number.tolist()[0]
    number = index[condition]
    row = number.tolist()[0]+2
    #print(row)
    ranges = 'D'+str(row)
    #print(ranges)
    update_sheet(ranges,list(['NO']),'Batch')
################################################################
####### Function to read dates for online batches ####

def get_dates_resi(start_date, weekday1, weekday2, no_sessions):
    wd = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'] 
    day1 = wd.index(weekday1)
    day2 = wd.index(weekday2)
    date1 = pd.to_datetime(start_date)
    day_start = date1.weekday()
    dates =[]
    duration = day1-day_start
    session_date1 = date1+timedelta(days=duration)
    dates.append(session_date1)
    count=1
    flag =0
    while count<= no_sessions:
        if flag == 0:
            duration = day2-day1
        else:
            duration =7
        session_date = dates[count-1]+timedelta(days=duration)
        dates.append(session_date)
        count+=1   
    return dates
#######################################################
def delet_row(sheet_name, row):
    sheet_del = data.get_sheet_by_name(sheet_name)
    sheet_del.delete_row(row,1)
    data.save(data_file)
################################################################################################

################# Function to delete the schedule #################
def deletion(batch,location,module):
    index = df_Schedule.index
    condition = (df_Schedule['Batch']==batch) & (df_Schedule['Location']==location) & (df_Schedule['Course']==module)
    number = index[condition]
    row = number.to_list()[0]+2
    print(row,type(row))
    data =df_Schedule.iloc[row-2:row-1]
    #print(data)
    faculty=data['Faculty'][row-2]
    sdate = data['Start Date'][row-2]
    edate = data['End Date'][row-2]
    module = data['Course'][row-2]
    stime = data['Start Time'][row-2]
    etime = data['End Time'][row-2]
    index1 = df_Calendar.index
    condition = (df_Calendar['Faculty']==faculty) & (df_Calendar['Date']>=sdate) & (df_Calendar['Date']<=edate) & (df_Calendar['Module']==module) & (df_Batch['Start_Time']=='9:30:00 AM')
    number1 = index1[condition]
    row1 = number1.to_list()[0]+2
    data1 = df_Calendar[condition]
    print(row1,data1.shape)
    delet_row("Schedule", row-1)
    for i in range(data1.shape[0]):
        delet_row('Calendar', row1-1)

######################################################################
################# Function to delete online schedule ##############
def deletion_online(batch,module):
    index = df_Online_Schedule.index
    condition = (df_Online_Schedule['Batch']==batch) &  (df_Online_Schedule['Course']==module)
    number = index[condition]
    row = number.to_list()[0]+2
    print(row,type(row))
    data =df_Online_Schedule.iloc[row-2:row-1]
    print(data)
    faculty=data['Faculty'][row-2]
    sdate = data['Start Date'][row-2]
    edate = data['End Date'][row-2]
    module = data['Course'][row-2]
    stime = data['Start Time'][row-2]
    etime = data['End Time'][row-2]
    index1 = df_Calendar.index
    condition = (df_Calendar['Faculty']==faculty) & (df_Calendar['Date']>=sdate) & (df_Calendar['Date']<=edate) & (df_Calendar['Module']==module) & (df_Calendar['Start_Time']=='8:00:00 PM')
    number1 = index1[condition]
    st.write(number1)
    row1 = number1.to_list()[0]+2
    data1 = df_Calendar[condition]
    print(row1,data1.shape)
    delet_row("Schedule_online", row-1)
    for i in range(data1.shape[0]):
        delet_row("Calendar", row1-1)
###################################################################
################# Update the data ########################
def update_batchdata(location,batch,sdate,edate,module,faculty,Remarks,type_session = 'Residency'):
    sdate = pd.to_datetime(sdate)
    edate = pd.to_datetime(edate)
    if location != 'Online':
        index = df_Batch.index
        condition = (df_Batch['Batch']==batch) & (df_Batch['Location']==location)
        number = index[condition]
        #print(number.tolist())
        #condition =number.tolist()[0]
        x=df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)].index
        s_date = df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Start_date'].iloc[0]
        old_faculty = df_Schedule[(df_Schedule['Batch']==batch) & (df_Schedule['Location']==location) & (df_Schedule['Course']==module)]['Faculty'].iloc[0]
        owner = df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Batch_Owner'].iloc[0]
        #latest_module = df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Latest_scheduled_Module'].iloc[0]
        if module == 'Case Study':
            active = 'NO'
        else:
            active = 'YES'
    #week = int(df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Week_of_Year'][0])+1
    #print(week)
        number = index[condition]
        row = number.tolist()[0]+2
        print(row)
        ranges = 'A'+str(row)
        print(ranges)
        update_sheet(ranges,list([batch, location, str(s_date),active, owner,module,str(sdate)]),'Batch')
        condition = (df_Schedule['Batch']==batch) & (df_Schedule['Location']==location) & (df_Schedule['Course']==module)
        index2 = df_Schedule.index
        number2 = index2[condition]
        row2 = number2.to_list()[0]+2
        ranges = 'A'+str(row2)
        ac = df_Batch[(df_Batch['Batch']==batch)&(df_Batch['Location']==location)]['Batch_Owner'].iloc[0]
        #Rem = df_Schedule[(df_Schedule['Batch']==batch)&(df_Schedule['Location']==location)&(df_Schedule['Course']==module)]['Remarks'].iloc[0]           
        update_sheet(ranges,list(['DSE-FT',location,str(sdate),str(edate),batch, module, faculty,ac,'9:30:00 AM','5:00:00 PM','Residency',2,7,Remarks]),'Schedule')
        index1 = df_Calendar.index
        print(old_faculty)
        condition = (df_Calendar['Faculty']==old_faculty) & (df_Calendar['Batch']==batch) & (df_Calendar['Location']==location) &(df_Calendar['Module']==module)
        number1 = index1[condition]
        print(number1)
        row1 = number1.to_list()[0]+2

        data1 = df_Calendar[condition]

        for i in range(data1.shape[0]):
            delet_row("Calendar", row1-1)
    elif location == 'Online':
        index = df_Onlinebatch.index
        condition = (df_Onlinebatch['Batch']==batch)
        number = index[condition]
        old_faculty = df_Online_Schedule[(df_Online_Schedule['Batch']==batch) & (df_Online_Schedule['Course']==module)]['Faculty'].iloc[0]
    #print(number.tolist())
    #condition =number.tolist()[0]
        s_date = df_Onlinebatch[df_Onlinebatch['Batch']==batch]['Start_date'].iloc[0]
        #latest_module = df_Onlinebatch[df_Onlinebatch['Batch']==batch]['Start_date'].iloc[0]
        if module == 'Case Study':
            active = 'NO'
        else:
            active = 'YES'
    #week = int(df_Batch[(df_Batch['Batch']==batch) & (df_Batch['Location']==location)]['Week_of_Year'][0])+1
    #print(week)
        number = index[condition]
        row = number.tolist()[0]+2
        print(row)
        ranges = 'Online_batch!A'+str(row)
        print(ranges)
        update_sheet(ranges,list([batch, str(s_date),active, module,str(sdate)]))
        index2 = df_Online_Schedule.index
        condition =(df_Online_Schedule['Batch']==batch)& (df_Online_Schedule['Course']==module) 
        number2 = index2[condition]
        row2 = number2.to_list()[0]+2
        ranges = 'A'+str(row2)
        if type_session == 'Residency':
            update_sheet(ranges,list(['DSE-Online',str(sdate),str(edate),batch,module,faculty,'8:00:00 PM', '10:00:00 PM',type_session,2,2,Remarks]),'Schedule_online')
        else:
            update_sheet(ranges,list(['DSE-Online',str(sdate),str(edate),batch,module,faculty,'9:30:00 AM', '1:30:00 PM',type_session,2,2,Remarks]),'Schedule_online')
        index1 = df_Calendar.index
        st.write(old_faculty)
        condition = (df_Calendar['Faculty']==old_faculty) & (df_Calendar['Batch']==batch) & (df_Calendar['Location']=='Online') &(df_Calendar['Module']==module)
        number1 = index1[condition]
        row1 = number1.to_list()[0]+2
        data1 = df_Calendar[condition]
        print(row1,data1.shape)
        delet_row("Schedule_online", row-1)
        for i in range(data1.shape[0]):
            delet_row("Calendar", row1-1) 

################# Function to get feedback #####################
def get_ratings(faculty,module):
    df_Rating['Faculty']=df_Rating['Faculty'].str.lower()
    df_Rating['Topic'] = df_Rating['Topic'].str.lower()
    faculty = faculty.lower()
    df_fac = pd.DataFrame(df_Rating[df_Rating['Faculty'].str.contains(faculty)])
    if module == 'SQL 1' or module == 'SQL 2':
        module = 'DBMS'
    elif module == 'Statistics':
        module = 'Stat'
    module = module.lower()
    df_fac = pd.DataFrame(df_fac[df_fac['Topic'].str.contains(module)])
    if df_fac.shape[0]==0:
        return '0'
    else:
        df_fac.sort_values(by='Session Date',ascending=False,inplace=True)
        df_fac = df_fac['Avg Ratings'].head(1)
        return df_fac.iloc[0]
#################################################################
###### Dates for the online batch lab sessions ###########
def get_dates_lab(start_date, weekday,no_sessions):
    dates = []
    date1 = pd.to_datetime(start_date)
    dates.append(date1)
    for i in range(no_sessions-1):
        date1 = date1+timedelta(days =7)
        dates.append(date1)
    return dates
#################################################################

################## Username Password Form #####################
def display_form():
   
          
    #user = st.text_input("User Name")
    '''   
    password1 = st.text_input("Password", type="password")
    submit3 =st.button("Login")
    if submit3:   
        if password1 == "Password@123":
            return True
        else:
            return False '''
    return True
######################################################################
#######################Restricted Access form ###################
def restrict():    
    st.error("Access Restricted")
    submit1 = st.button("Ok")
    if submit1:
        #st.empty()
        st.stop()
###############################################################
################# Function to calculate skills in Yes/No format ###############
def skill_value(skill):
    print("In skill vale")
    if skill:
        return 'Yes'
    else:
        return 'No'
##############################################################

############ Function to add new FT batch ############
def add_ft_batch():
    batch_name = st.text_input("Batch Name")
    location = st.selectbox("City",['Banglore','Chennai','Gurgaon','Hyderabad','Mumbai','Pune'])
    start_date = st.date_input("Start Date")
    is_active = 'YES'
    faculty = list(df_Faculty['Faculty Name'])
    Batch_Owner = st.selectbox("Batch Owner",faculty)
    Latest_module = "Starter"
    Latest_schedule_date = start_date
    week = 0
    Submit = st.button("Submit")
    if Submit:
        write_to_sheet('Batch',list(batch_name,location,start_date,is_active,Batch_Owner,Latest_module,Latest_schedule_date,week))

#########################################################################################
############ Function to add new Online batch ############
def add_online_batch():
    batch_name = st.text_input("Batch Name")
    start_date = st.date_input("Start Date")
    is_active = 'YES'
    Latest_module = "Starter"
    Latest_schedule_date = start_date
    Submit = st.button("Submit")
    if Submit:
        write_to_sheet('Online_Batch',list(batch_name,start_date,is_active,Latest_module,Latest_schedule_date))
############# Function to add a new Faculty ##############
def add_faculty():
    #caching.clear_cache()  
    fname = st.text_input("Faculty Name")
    st.write("Choose Skills")
        #module_liform1 = liform1(df_Modules['Module Name'])

    list_skills = ['ITP','NPV', 'SQL1', 'SQL2', 'EDA',  'Statistics','SLR', 'SLC',  
        'USL',  'Python_LI',    'ITP_LI',   'NPV_LI',   'SQL_LI',   'SQL1_LI',  'SQL2_LI',  
        'EDA_STAT_LI',  'EDA_LI',   'STAT_LI',  'SLR_LI',   'SLC_USL_LI',   'SLC_LI',   'USL_LI',
        'ITP_Lab',  'NPV_Lab'   ,'SQL1_Lab',    'SQL2_Lab', 'EDA_Lab',  
        'Statistics_Lab',   'SLR_Lab',  'SLC_Lab',  'USL_Lab',  'Case_study','Placement_Bootcamp']
        #print(len(liform1_skills))
    list_skill_value=[0 for i in range(len(list_skills))]
        #c1,c2,c3 = form1.columns(3)
        #with c1:
    list_skill_value[0] = st.checkbox("ITP")
    list_skill_value[1] = st.checkbox("NPV")
    list_skill_value[2] = st.checkbox("SQL1")
    list_skill_value[3] = st.checkbox("SQL2")
    list_skill_value[4] = st.checkbox("EDA")
    list_skill_value[31]= st.checkbox("Case form1udy")
    list_skill_value[32]= st.checkbox("Bootcamp")
    #with c2:
    list_skill_value[5]= st.checkbox("Statistics")
    list_skill_value[6] = st.checkbox("SLR")
    list_skill_value[7]= st.checkbox("SLC")
    list_skill_value[8]= st.checkbox("USL")
            #with c3:
    list_skill_value[9] = st.checkbox("Python_LI")
    if list_skill_value[9]:
        list_skill_value[10] = True # ITP LI
        list_skill_value[11] = True #NPV LI
    list_skill_value[12]=st.checkbox("SQL_LI")
    if list_skill_value[12]:
        list_skill_value[13]=True #SQL 1 LI
        list_skill_value[14] = True # SQL2 LI

    list_skill_value[15] = st.checkbox("EDA_STAT_LI")
    if list_skill_value[15]:
        list_skill_value[16]=True #EDA LI
        list_skill_value[17] = True #form1at LI
    list_skill_value[18] = st.checkbox("SLR_LI")
    list_skill_value[19]=st.checkbox("SLC_USL_LI")
    if list_skill_value[19]:
        list_skill_value[20]= True #SLC LI
        list_skill_value[21] = True  #USL LI

        #c4,c5 = form1.columns(2)
        #with c4:
    list_skill_value[22] = st.checkbox("ITP Lab")
    list_skill_value[23]= st.checkbox("NPV Lab")
    list_skill_value[24] = st.checkbox("SQL1 Lab")
    list_skill_value[25] = st.checkbox("SQL2 Lab")
    #with c5:
    list_skill_value[26] = st.checkbox("EDA Lab")
    list_skill_value[27] = st.checkbox("Statistics Lab")
    list_skill_value[28] = st.checkbox("SLR Lab")
    list_skill_value[29] = st.checkbox("SLC Lab")
    list_skill_value[30] = st.checkbox("USL Lab")
    print(list_skill_value)
    finternal = st.radio("Internal/External",["Internal","External"])
    flocation = st.selectbox('City',('Banglore','Chennai','Gurgaon','Hyderabad','Mumbai','Pune','Other'))

    print(finternal,flocation)
    if flocation == 'Other':
        st.text('Write the city name in case you choose Other')
        flocation = st.text_input("City")
        #wh form1:
    submit4 = st.button("Submit")
    print(submit4)
    if submit4:
        print("submit to process")
        for i in range(len(list_skill_value)):
            list_skill_value[i] = skill_value(list_skill_value[i])
                
    # SQL1  SQL 2   EDA form1atiform1ics    SLR SLC USL Python LI   ITP LI  NPV LI  SQL LI  SQL 1 LI    SQL 2 LI    EDA_form1AT LI  EDA LI  form1AT LI  SLR LI  SLC_USL LI  SLC LI  USL LI  ITP Lab NPV Lab SQL1 Lab    SQL 2 Lab   EDA Lab form1atiform1ics Lab    SLR Lab SLC Lab USL Lab Case form1udy   Placement Bootcamp
        print("Ready to write the data for Faculty",fname,finternal,flocation,list_skill_value)               
        write_to_sheet('Faculty',list([fname,list_skill_value[0],list_skill_value[1],list_skill_value[2],list_skill_value[3],
        list_skill_value[4],list_skill_value[5],list_skill_value[6],list_skill_value[7],list_skill_value[8],    
        list_skill_value[9],list_skill_value[10],list_skill_value[11],list_skill_value[12],
        list_skill_value[13],list_skill_value[14],list_skill_value[15],list_skill_value[16],
        list_skill_value[17],list_skill_value[18],list_skill_value[19],
        list_skill_value[20],list_skill_value[21],list_skill_value[22], list_skill_value[23],
        list_skill_value[24],list_skill_value[25],list_skill_value[26],list_skill_value[27],
        list_skill_value[28],list_skill_value[29],list_skill_value[30],list_skill_value[31],list_skill_value[32],
                finternal,  flocation]))
######### GUI Code #####################
### Choosing Between Full time and Online Schedule
st.title("Great Learning")
st.header("Faculty Planner")
flag = False
correct = True
application =''
user= ''
application  = st.sidebar.radio("Choose Application",['View Existing','Create New Schedule','Create New Entry','Update','Delete'])
if application == 'Login':
    admin = st.sidebar.radio("User",['Viewer','Admin'])      
#submits = st.button("Submit")
#if submits:
    #user = admin
    
if correct:
    if application == 'Delete':
        flag = display_form()
    #with st.form("Edit Form", clear_on_submit=False):
        if flag:
            st.write('Choose the batch and schedule details')
            batch_type = st.radio('Batch Type',['FT','Online'])
            if batch_type == 'FT':
                location = st.selectbox('City',('Banglore','Chennai','Gurgaon','Hyderabad','Mumbai','Pune'))
                df_active = all_active(location)
        #print(list(df_active.Batch))
                batch = st.selectbox('Select Batch',list(df_active.Batch))
                from_date = st.date_input("From Date")
                to_date = st.date_input("To Date")
                df = read_schedule(location,batch,from_date,to_date)
                st.write(df)
        
        #if operation == 'Delete':
                st.write("Select Module to delete")
                module_list = list(df_Modules['Module Name'])
                module = st.selectbox('Module Name',module_list)
                con = st.button("Confirm deletion",101)
                if con:
                    deletion(batch,location,module)
        
        #st.experimental_rerun()
            elif batch_type == 'Online':
                batches = all_onlineactive()
                batch = st.selectbox('Batch',list(batches.Batch))
                from_date = st.date_input("From Date")
                to_date = st.date_input("To Date")
                df = read_online_schedule(batch,from_date,to_date)
                st.write(df)
        #operation = st.radio("Edit Type",['Delete','Update'])
        #if operation == 'Delete':
                st.write("Select Module to delete")
                module_list = list(df_Modules['Module Name'])
                module = st.selectbox('Module Name',module_list)
                con1 = st.button("Confirm deletion",103)
                if con1:
                    deletion_online(batch,module)

    
    if application=='Create New Schedule':
        flag = display_form()
        if flag:
            batch_type = st.sidebar.radio( "Schedule", ("Full-Time", "Online"))

            print(batch_type)
            if batch_type == 'Full-Time':
                location = st.selectbox('City',('Banglore','Chennai','Gurgaon','Hyderabad','Mumbai','Pune'))
                df_active = all_active(location)
        #print(list(df_active.Batch))
                batch = st.selectbox('Select Batch',list(df_active.Batch))
                views = st.button("View Existing Schedule",201)
                if views:
                    df = read_schedule(location,batch)
                    st.write(df)
                choice = st.radio('Do you want to schedule?',('Yes','No'))
                if choice == 'Yes':
                    module_list = list(df_Modules['Module Name'])
                    print(module_list)
                    module = st.selectbox('Module Name',module_list)
                    df_all = search_faculty(module,location)
                    print("*****DF_all******",df_all,sep='\n')
            #df_all.sort_values('Weight', axis=0, ascending=False, inplace=True )
                    start_date = st.date_input('Start Date',min_value=datetime.date(2021, 1, 1))
                    end_date = st.date_input('End Date',min_value=datetime.date(2021, 1, 1))
                    av=[]
                    ratings=[]
                    for fac in df_all['Name']:
                        ratings.append(get_ratings(fac,module))
                        av.append(check_availability(fac,start_date, end_date, '9:30:00 AM', '5:00:00 PM','FT'))
                    df_all['Available'] = av
                    df_all['Ratings']=ratings
                    print("**************",df_all)
                    df_all.sort_values(['Available','weight'], axis=0, ascending=False, inplace=True )
                    st.write(df_all)
                    st.write('Proceed to schedule')
                    faculty = st.selectbox('Faculty',df_all['Name'])
                    available = check_availability(faculty,start_date, end_date, '9:30:00 AM', '5:00:00 PM','FT')
                    if available:
                        st.write('Faculty is available')
                        Remarks = st.selectbox('Remarks',['Revision','Remedial','Prep','Extra','NA'])
                        st.write('Press Submit to confirm the schedule')
                        save = st.button('Submit')
                        if save:
                            ac = df_Batch[(df_Batch['Batch']==batch)&(df_Batch['Location']==location)]['Batch_Owner'].iloc[0]
                            print(ac)
                            write_to_sheet('Schedule',list(['DSE-FT',location,str(start_date), str(end_date),batch,module,faculty,ac,'9:30:00 AM', '5:00:00 PM','Residency',7,7,Remarks])) 
                            dates=pd.date_range(start= start_date, end=end_date)
                            for d in dates:
                                write_to_sheet('Calendar',list([faculty,str(d),'9:30:00 AM', '5:00:00 PM',module,location,batch]))
                            set_batch(batch,location,str(start_date),module)
                    
                    #final=st.radio('Do you want confirm?',('Yes','No'))
                    #if final == 'Yes':
                        ## write the code to write to sheet calendar
                        #ac = df_Batch[(df_Batch['Batch']==batch)&(df_Batch['Location']==location)]['Batch_Owner']
                            print()
                        #write_to_sheet('Schedule',list(['DSE-FT',location,str(start_date), str(end_date),batch,status[2],faculty,faculty,'9:30:00 AM', '5:00:00 PM','Residency',7,7]),SPREADSHEET_ID) 
                        #dates=pd.date_range(start= start_date, end=end_date)
                        #for d in dates:
                            #write_to_sheet('Calendar',list([faculty,str(d),'9:30:00 AM', '5:00:00 PM']),SPREADSHEET_ID)'''
                    else:
                        st.write('Faculty not avaialable')
                         
            
            else:
        
                batches = all_onlineactive()
                batch = st.selectbox('Batch',list(batches.Batch))
                from_date = st.date_input("From Date")
                to_date = st.date_input("To Date")
                df = read_online_schedule(batch,from_date,to_date)
                st.write(df)
                choice = st.radio('Do you want to schedule?',('Yes','No'))
                if choice == 'Yes':
                    module_list = list(df_Modules['Module Name'])
                    module = st.selectbox('Module Name',module_list)
 
                    df_all = search_faculty(module,'Online')
                    df_all.sort_values('weight', axis=0, ascending=False, inplace=True )
                    start_date = st.date_input('Start Date',min_value=datetime.date(2021, 1, 1))
                    session_count = int(st.text_input("Number of sessions",value="0"))
                    type_session = st.radio("Type of Session",['Residency','Lab'])
                    st.write("Select days of week")
                    days = st.multiselect('Week Days',['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'])
                    if type_session == 'Residency':
                        resi_dates=get_dates_resi(start_date,days[0],days[1],session_count)
                    elif type_session == 'Lab':
                        resi_dates=get_dates_lab(start_date,days[0],session_count)
                    av=[]
                    ratings=[]
                    for fac in df_all['Name']:
                        ratings.append(get_ratings(fac,module))
                        result = True
                        for d in resi_dates:
                            if type_session == ' Residency':
                                result = check_availability(fac,d,d, '8:00:00 PM', '10:00:00 PM','FT')
                            else:
                                result = check_availability(fac,d,d,'9:30:00 AM','1:30:00 PM', 'FT')
                            if result == 'False':
                                break
                        av.append(result)
                    df_all['Available'] = av
                    df_all['Ratings']=ratings

                    df_all.sort_values(['Available','weight'], axis=0, ascending=False, inplace=True )
                    st.write(df_all)
                    st.write('Proceed to schedule')
                    faculty = st.selectbox('Faculty',df_all['Name'])
                    for d in resi_dates:
                        if type_session == 'Residency':
                            result = check_availability(faculty,d,d, '8:00:00 PM', '10:00:00 PM','FT')
                        else:
                            result = check_availability(faculty,d,d,'9:30:00 AM','1:30:00 PM','FT')
                        if result == 'False':
                            st.write("Faculty not available")
                            break
                    else:
                        st.write("Faculty is available")
                        st.write('Press Submit to confirm the schedule')
                        Remarks = st.selectbox('Remarks',['Revision','Remedial','Prep','Extra','NA'])
                        save = st.button('Submit')
                        if save:
                        ## To do work
                            if type_session == 'Residency':
                                write_to_sheet('Schedule_online',list(['DSE-Online',str(resi_dates[0]), str(resi_dates[-1]),batch,module,faculty,'8:00:00 PM', '10:00:00 PM',type_session,2,2,Remarks])) 
                            else:
                                write_to_sheet('Schedule_online',list(['DSE-Online',str(resi_dates[0]), str(resi_dates[-1]),batch,module,faculty,'9:30:00 AM', '1:30:00 PM',type_session,2,2,Remarks])) 
                            for d in resi_dates:
                                if type_session == 'Residency':
                                    write_to_sheet('Calendar',list([faculty,str(d),'8:00:00 PM', '10:00:00 PM',module,'Online',batch]))
                                else:
                                    write_to_sheet('Calendar',list([faculty,str(d),'9:30:00 AM', '1:30:00 PM',module,'Online',batch]))
                            set_onlinebatch(batch,str(start_date),module)
       
   # else:
    #    restrict()
    if application == 'Create New Entry':
        flag = display_form()
        if flag:
        
            profile = st.sidebar.radio("New",["Faculty Profile","FT Batch Entry","Online Batch Entry"])
            if profile == 'Faculty Profile':
                add_faculty()
            
                        #flag=False
            elif profile == "FT Batch Entry":
                add_ft_batch()
            elif profile == "Online Batch Entry":
                add_online_batch()
                               
    if application == 'Update':
        flag = display_form()
        if flag:
    
            btype = st.sidebar.radio("Batch Type",('Full-Time','Online'))
            if btype == 'Full-Time':
        
                location = st.selectbox('City',('Banglore','Chennai','Gurgaon','Hyderabad','Mumbai','Pune'))
                df_active = all_active(location)
            #print(list(df_active.Batch))
                batch = st.selectbox('Select Batch',list(df_active.Batch))
                df = read_schedule(location,batch)
                st.write(df)
                module_list = list(df_Modules['Module Name'])
            #print(module_list)
                st.write("Enter new data for update")
                module = st.selectbox('Module to update',module_list)
                df_all = search_faculty(module,location)
            #print("*****DF_all******",df_all,sep='\n')
            #df_all.sort_values('Weight', axis=0, ascending=False, inplace=True )
                start_date = st.date_input('Start Date',min_value=datetime.date(2021, 1, 1))
                end_date = st.date_input('End Date',min_value=datetime.date(2021, 1, 1))
                av=[]
                ratings=[]
                for fac in df_all['Name']:
                    ratings.append(get_ratings(fac,module))
                    av.append(check_availability(fac,start_date, end_date, '9:30:00 AM', '5:00:00 PM','FT'))
                df_all['Available'] = av
                df_all['Ratings']=ratings
            #print("**************",df_all)
                df_all.sort_values(['Available','weight'], axis=0, ascending=False, inplace=True )
                st.write(df_all)
                st.write('Proceed to update')
                faculty = st.selectbox('Faculty',df_all['Name'])
                df_temp = df_Schedule[(df_Schedule['Location']==location) & (df_Schedule['Batch']==batch) & (df_Schedule['Course']==module)]
                if df_temp.shape[0]!=0:
                    fac = df_temp['Faculty'].iloc[0]

                    if faculty == fac:
                        available=True
                else:
                    available = check_availability(faculty,start_date, end_date, '9:30:00 AM', '5:00:00 PM','FT')
                if available:
                    st.write('Faculty is available')
                    Remarks = st.selectbox('Remarks',['Residency','Revision','Remedial','Prep','Extra','NA'])
                    st.write('Press Submit to confirm the schedule')
                    save = st.button('Submit')
                    if save:
                        ac = df_Batch[(df_Batch['Batch']==batch)&(df_Batch['Location']==location)]['Batch_Owner'].iloc[0]
                    #print(ac)
                #write_to_sheet('Schedule',list(['DSE-FT',location,str(start_date), str(end_date),batch,module,faculty,ac,'9:30:00 AM', '5:00:00 PM','Residency',7,7,Remarks]),SPREADSHEET_ID) 
                        update_batchdata(location,batch,start_date,end_date,module,faculty,Remarks)
                        dates=pd.date_range(start= start_date, end=end_date)
                        for d in dates:
                            write_to_sheet('Calendar',list([faculty,str(d),'9:30:00 AM', '5:00:00 PM',module,location,batch]))
                        set_batch(batch,location,str(start_date),module)
       
            elif btype == 'Online':
                batches = all_onlineactive()
                batch = st.selectbox('Batch',list(batches.Batch))
                df = read_online_schedule(batch)
                st.write(df)
                module_list = list(df_Modules['Module Name'])
            #print(module_list)
                st.write("Enter new data for update")
                module = st.selectbox('Module to update',module_list)
        #df_all = search_faculty(module,location)
                df_all = search_faculty(module,'Online')
        #df_all.sort_values('weight', axis=0, ascending=False, inplace=True )
                start_date = st.date_input('Start Date',min_value=datetime.date(2021, 1, 1))
                session_count = int(st.text_input("Number of sessions",value="0"))
                type_session = st.radio("Type of Session",['Residency','Lab'])
                st.write("Select days of week")
                days = st.multiselect('Week Days',['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'])
                if type_session == 'Residency':
                    resi_dates=get_dates_resi(start_date,days[0],days[1],session_count)
                elif type_session == 'Lab':
                    resi_dates=get_dates_lab(start_date,days[0],session_count)
                av=[]
                ratings=[]
                for fac in df_all['Name']:
                    ratings.append(get_ratings(fac,module))
                    result = True
                    for d in resi_dates:
                        result = check_availability(fac,d,d, '8:00:00 PM', '10:00:00 PM','FT')
                        if result == 'False':
                            break
                    av.append(result)
                df_all['Available'] = av
                df_all['Ratings']=ratings

                df_all.sort_values(['Available','weight'], axis=0, ascending=False, inplace=True )
                st.write(df_all)
                st.write('Proceed to update')
                faculty = st.selectbox('Faculty',df_all['Name'])
                result = True
                if faculty == df_Online_Schedule[(df_Online_Schedule['Batch']==batch) & (df_Online_Schedule['Course']==module)]['Faculty'].iloc[0]:
                    result = True
                else:
                    for d in resi_dates:
                        if type_session == 'Residency':
                            result = check_availability(faculty,d,d, '8:00:00 PM', '10:00:00 PM','FT')
                        else:
                            result = check_availability(faculty,d,d,'9:30:00 AM','1:30:00 PM','FT')
                        if result == 'False':
                            st.write("Faculty not available")
                            break
                    else:
                        st.write("Faculty is available")
                        st.write('Press Submit to confirm the schedule')
                        Remarks = st.selectbox('Remarks',['Revision','Remedial','Prep','Extra','NA'])
                        save = st.button('Submit')
                        if save:
                    ## To do work
                            if type_session == 'Residency':
                                update_batchdata('Online',batch,resi_dates[0],resi_dates[-1],module,faculty,Remarks,'Residency') 
                            else:
                                update_batchdata('Online',batch,resi_dates[0],resi_dates[-1],module,faculty,Remarks,'Lab') 
                            for d in resi_dates:
                                if type_session == 'Residency':
                                    write_to_sheet('Calendar',list([faculty,str(d),'8:00:00 PM', '10:00:00 PM',module,'Online',batch]))
                                else:
                                    write_to_sheet('Calendar',list([faculty,str(d),'9:30:00 AM', '1:30:00 PM',module,'Online',batch]))
                            set_onlinebatch(batch,str(start_date),module)
  

if application=='View Existing':
    view =st.sidebar.radio("Select View",("Faculty View","Batch View"))
    if view=='Batch View':
        batch_type = st.radio('Type of batch',['FT','Online'])
        if batch_type == 'FT':
            location = st.selectbox('City',('Banglore','Chennai','Gurgaon','Hyderabad','Mumbai','Pune'))
            df_active = all_active(location)
        #print(list(df_active.Batch))
            batch = st.selectbox('Select Batch',list(df_active.Batch))
            from_date = st.date_input("From Date")
            to_date = st.date_input("To Date")
            df = read_schedule(location,batch,from_date,to_date)
            st.write(df)
        elif batch_type == 'Online':
            batches = all_onlineactive()
            batch = st.selectbox('Batch',list(batches.Batch))
            from_date = st.date_input("From Date")
            to_date = st.date_input("To Date")
            df = read_online_schedule(batch,from_date,to_date)
            st.write(df)
    if view=='Faculty View':
        names = list(df_Faculty['Faculty Name'])
        fac = st.selectbox("Faculty",names)
        from_date = st.date_input("From Date")
        to_date = st.date_input("To Date")
        df = read_faculty_schedule(fac,from_date,to_date)
        st.write(df)    
    
    
    

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
                

