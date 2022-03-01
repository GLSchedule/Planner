import pandas as pd
import streamlit as st
import datetime
import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow,Flow
from google.auth.transport.requests import Request
#from datetime import datetime
from datetime import timedelta
## Required for writing to goolge sheet
from pprint import pprint
from googleapiclient import discovery


## Google Sheet Credentials
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# Original Spreadsheetid
#SPREADSHEET_ID = '1A2spx394hFZEhArmSc0rxfgYddoQjs97QW-EiCF0U_4'
SPREADSHEET_ID = '1V3jT3A3g_8TnIXv5e8ZdLlY8IoPK3ubV9_UmUu8V9ws'
## Function to open a sheet and load the data in dataframe
def main(SAMPLE_SPREADSHEET_ID_input,SAMPLE_RANGE_NAME):
    #values_input, service
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('google1.json', SCOPES) # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
                                range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])
   
    if not values_input :
        print('No data found.')
        return -1
    return values_input
#############################################

## Faculty sheet Load
SHEET_NAME ='Faculty'
SAMPLE_RANGE = 'Faculty'
values_input = main(SPREADSHEET_ID,SAMPLE_RANGE)
df_Faculty=pd.DataFrame(values_input[1:], columns=values_input[0])
# st.write(df_Faculty)
#############################################

## Schedule sheet Load

values_input = main(SPREADSHEET_ID,'Schedule')
df_Schedule=pd.DataFrame(values_input[1:], columns=values_input[0])
#st.write(df_Schedule)

## Converting "Start Date" and "end Date" columns to DateTime datatype 

df_Schedule['Start Date'] = pd.to_datetime(df_Schedule['Start Date'])
df_Schedule['End Date'] = pd.to_datetime(df_Schedule['End Date'])
#st.write(df_Schedule)
##############################################

## Calendar sheet Load

values_input = main(SPREADSHEET_ID,'Calendar')
df_Calendar=pd.DataFrame(values_input[1:], columns=values_input[0])

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

df_Schedule['Start Time'] = df_Schedule['Start Time'].apply(time_convert)
df_Schedule['End Time'] = df_Schedule['End Time'].apply(time_convert)
df_Calendar['Start_Time'] = df_Calendar['Start_Time'].apply(time_convert)
df_Calendar['End_Time'] = df_Calendar['End_Time'].apply(time_convert)


#st.write(df_Calendar)
### Reading Ratings ########### 
######## Ratings fetched from Metabase--> Copy paste mannually

values_input = main(SPREADSHEET_ID,'Rating')

df_Rating=pd.DataFrame(values_input[1:], columns=values_input[0])
#### Reading batch data ###########
values_input = main(SPREADSHEET_ID,'Batch')

df_Batch=pd.DataFrame(values_input[1:], columns=values_input[0])
df_Batch['Latest_Scheduled_Date'] = pd.to_datetime(df_Batch['Latest_Scheduled_Date'])

####### Reading Module sequence #########
values_input = main(SPREADSHEET_ID,'Modules')

df_Modules=pd.DataFrame(values_input[1:], columns=values_input[0])

## Reading Weight matrix ############
values_input = main(SPREADSHEET_ID,'Weight')

df_Weight=pd.DataFrame(values_input[1:], columns=values_input[0])
##########################################################

values_input = main(SPREADSHEET_ID,'Schedule_online')

df_Online_Schedule=pd.DataFrame(values_input[1:], columns=values_input[0])
df_Online_Schedule['Start Date'] = pd.to_datetime(df_Online_Schedule['Start Date'])
df_Online_Schedule['End Time'] = df_Online_Schedule['End Time'].apply(time_convert)
df_Online_Schedule['Start Time'] = df_Online_Schedule['Start Time'].apply(time_convert)
#############################################################
############# Reading online bathes ############
values_input = main(SPREADSHEET_ID,'Online_batch')

df_Onlinebatch=pd.DataFrame(values_input[1:], columns=values_input[0])
df_Onlinebatch['Latest_Scheduled_Week'] = pd.to_datetime(df_Onlinebatch['Latest_Scheduled_Week'])
########### Function to write to Google sheet #########

###
def write_to_sheet(place,data,sheet_id):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    credentials = None

    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'google1.json', SCOPES) # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
#service = discovery.build('sheets', 'v4', credentials=credentials)
    spreadsheet_id = sheet_id  # TODO: Update placeholder value.

# The A1 notation of a range to search for a logical table of data.
# Values will be appended after the last row of the table.
    range_ = place  # TODO: Update placeholder value.

# How the input data should be interpreted.
    value_input_option = 'RAW'  # TODO: Update placeholder value.

# How the input data should be inserted.
    insert_data_option = 'INSERT_ROWS'  # TODO: Update placeholder value.

    value_range_body = {  "values": [data]}
    # TODO: Add desired entries to the request body.

    request = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_, valueInputOption=value_input_option, insertDataOption=insert_data_option, body=value_range_body)
    response = request.execute()

# TODO: Change code below to process the `response` dict:
    pprint(response)
###########################################################################

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
    return active_batches.loc[:,['Batch']]
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
    print(list_faculty)
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
        print(dict_faculty,inter,sep="|")
    
    for x in dict_faculty.values():
        list_wei.append(x)
    d=dict(Name=list_fac,loc=list_loc,inernal=list_int,weight=list_wei)
    print(d)
    df_try = pd.DataFrame.from_dict(d)
    print(df_try)
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
    df_faculty_schedule = df_Schedule[df_Calendar['Faculty']==faculty]
    df_faculty_schedule = df_faculty_schedule[(df_faculty_schedule['Start Date'].dt.date>=sdate) & (df_faculty_schedule['Start Date'].dt.date<=ldate)]
    return df_faculty_schedule
################## Function to update the sheet ########################
# TODO: Change placeholder below to generate authentication credentials. See
# https://developers.google.com/sheets/quickstart/python#step_3_set_up_the_sample
#
# Authorize using one of the following scopes:
#     'https://www.googleapis.com/auth/drive'
#     'https://www.googleapis.com/auth/drive.file'
#     'https://www.googleapis.com/auth/spreadsheets'
def update_sheet(place,data,sheet_id):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    credentials = None

    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'google1.json', SCOPES) # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
#service = discovery.build('sheets', 'v4', credentials=credentials)
    spreadsheet_id = sheet_id  # TODO: Update placeholder value.

# The A1 notation of a range to search for a logical table of data.
# Values will be appended after the last row of the table.
    range_ = place  # TODO: Update placeholder value.

# How the input data should be interpreted.
    value_input_option = 'RAW'  # TODO: Update placeholder value.

# How the input data should be inserted.
    insert_data_option = 'INSERT_ROWS'  # TODO: Update placeholder value.

    value_range_body = {  "values": [data]}
    # TODO: Add desired entries to the request body.
    request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_, valueInputOption=value_input_option, body=value_range_body)
    response = request.execute()
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
    print(row)
    ranges = 'Batch!A'+str(row)
    print(ranges)
    update_sheet(ranges,list([batch, location, str(s_date),active, owner,latest_module,str(latest_date)]),SPREADSHEET_ID)
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
    print(row)
    ranges = 'Online_batch!A'+str(row)
    print(ranges)
    update_sheet(ranges,list([batch, str(s_date),active, latest_module,str(latest_date)]),SPREADSHEET_ID)
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
    print(row)
    ranges = 'Batch!D'+str(row)
    print(ranges)
    update_sheet(ranges,list(['NO']),SPREADSHEET_ID)
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
################ Function to delete a row from google sheet #########
def delet_row(spreadsheet_id, sheet_id, index):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    credentials = None

    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'google1.json', SCOPES) # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    request_body={
      "requests": [
    {
      "deleteDimension": {
        "range": {
          "sheetId": sheet_id,
          "dimension": "ROWS",
          "startIndex": index,
          "endIndex": index+1 }       }     },
      ], }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=request_body).execute()
################################################################################################

################# Function to delete the schedule #################
def deletion(spreadsheet_id,batch,location,module):
    index = df_Schedule.index
    condition = (df_Schedule['Batch']==batch) & (df_Schedule['Location']==location) & (df_Schedule['Course']==module)
    number = index[condition]
    row = number.to_list()[0]+2
    print(row,type(row))
    data =df_Schedule.iloc[row-2:row-1]
    print(data)
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
    delet_row(spreadsheet_id, 0, row-1)
    for i in range(data1.shape[0]):
        delet_row(spreadsheet_id, 433910658, row1-1)

######################################################################
################# Function to delete online schedule ##############
def deletion_online(spreadsheet_id,batch,module):
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
    delet_row(spreadsheet_id, 1227198579, row-1)
    for i in range(data1.shape[0]):
        delet_row(spreadsheet_id, 433910658, row1-1)
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
        ranges = 'Batch!A'+str(row)
        print(ranges)
        update_sheet(ranges,list([batch, location, str(s_date),active, owner,module,str(sdate)]),SPREADSHEET_ID)
        condition = (df_Schedule['Batch']==batch) & (df_Schedule['Location']==location) & (df_Schedule['Course']==module)
        index2 = df_Schedule.index
        number2 = index2[condition]
        row2 = number2.to_list()[0]+2
        ranges = 'Schedule!A'+str(row2)
        ac = df_Batch[(df_Batch['Batch']==batch)&(df_Batch['Location']==location)]['Batch_Owner'].iloc[0]
        #Rem = df_Schedule[(df_Schedule['Batch']==batch)&(df_Schedule['Location']==location)&(df_Schedule['Course']==module)]['Remarks'].iloc[0]           
        update_sheet(ranges,list(['DSE-FT',location,str(sdate),str(edate),batch, module, faculty,ac,'9:30:00 AM','5:00:00 PM','Residency',2,7,Remarks]),SPREADSHEET_ID)
        index1 = df_Calendar.index
        print(old_faculty)
        condition = (df_Calendar['Faculty']==old_faculty) & (df_Calendar['Batch']==batch) & (df_Calendar['Location']==location) &(df_Calendar['Module']==module)
        number1 = index1[condition]
        print(number1)
        row1 = number1.to_list()[0]+2

        data1 = df_Calendar[condition]

        for i in range(data1.shape[0]):
            delet_row(SPREADSHEET_ID, 433910658, row1-1)
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
        update_sheet(ranges,list([batch, str(s_date),active, module,str(sdate)]),SPREADSHEET_ID)
        index2 = df_Online_Schedule.index
        condition =(df_Online_Schedule['Batch']==batch)& (df_Online_Schedule['Course']==module) 
        number2 = index2[condition]
        row2 = number2.to_list()[0]+2
        ranges = 'Schedule_online!A'+str(row2)
        if type_session == 'Residency':
            update_sheet(ranges,list(['DSE-Online',str(sdate),str(edate),batch,module,faculty,'8:00:00 PM', '10:00:00 PM',type_session,2,2,Remarks]),SPREADSHEET_ID)
        else:
            update_sheet(ranges,list(['DSE-Online',str(sdate),str(edate),batch,module,faculty,'9:30:00 AM', '1:30:00 PM',type_session,2,2,Remarks]),SPREADSHEET_ID)
        index1 = df_Calendar.index
        st.write(old_faculty)
        condition = (df_Calendar['Faculty']==old_faculty) & (df_Calendar['Batch']==batch) & (df_Calendar['Location']=='Online') &(df_Calendar['Module']==module)
        number1 = index1[condition]
        row1 = number1.to_list()[0]+2
        data1 = df_Calendar[condition]
        print(row1,data1.shape)
        #delet_row(spreadsheet_id, 1227198579, row-1)
        for i in range(data1.shape[0]):
            delet_row(SPREADSHEET_ID, 433910658, row1-1) 

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
######### GUI Code #####################
### Choosing Between Full time and Online Schedule
st.header("Faculty Planner")
application = st.sidebar.radio("Choose Application",('View Existing','Create New','Update','Delete'))
if application == 'Delete':
    #with st.form("Edit Form", clear_on_submit=False):
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
            deletion(SPREADSHEET_ID,batch,location,module)
        
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
            deletion_online(SPREADSHEET_ID,batch,module)
       
if application=='Create New':
    batch_type = st.sidebar.radio( "Type of Batch", ("Full-Time", "Online"))

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
                    write_to_sheet('Schedule',list(['DSE-FT',location,str(start_date), str(end_date),batch,module,faculty,ac,'9:30:00 AM', '5:00:00 PM','Residency',7,7,Remarks]),SPREADSHEET_ID) 
                    dates=pd.date_range(start= start_date, end=end_date)
                    for d in dates:
                        write_to_sheet('Calendar',list([faculty,str(d),'9:30:00 AM', '5:00:00 PM',module,location,batch]),SPREADSHEET_ID)
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
                        write_to_sheet('Schedule_online',list(['DSE-Online',str(resi_dates[0]), str(resi_dates[-1]),batch,module,faculty,'8:00:00 PM', '10:00:00 PM',type_session,2,2,Remarks]),SPREADSHEET_ID) 
                    else:
                         write_to_sheet('Schedule_online',list(['DSE-Online',str(resi_dates[0]), str(resi_dates[-1]),batch,module,faculty,'9:30:00 AM', '1:30:00 PM',type_session,2,2,Remarks]),SPREADSHEET_ID) 
                    for d in resi_dates:
                        if type_session == 'Residency':
                            write_to_sheet('Calendar',list([faculty,str(d),'8:00:00 PM', '10:00:00 PM',module,'Online',batch]),SPREADSHEET_ID)
                        else:
                            write_to_sheet('Calendar',list([faculty,str(d),'9:30:00 AM', '1:30:00 PM',module,'Online',batch]),SPREADSHEET_ID)
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

            
if application == 'Update':
    btype = st.sidebar.radio("Batch Type",('Full-Time','Online'))
    if btype == 'Full-Time':
        
        location = st.selectbox('City',('Banglore','Chennai','Gurgaon','Hyderabad','Mumbai','Pune'))
        df_active = all_active(location)
            #print(list(df_active.Batch))
        batch = st.selectbox('Select Batch',list(df_active.Batch))
        df = read_schedule(location,batch)
        st.write(df)
        module_list = list(df_Modules['Module Name'])
        print(module_list)
        st.write("Enter new data for update")
        module = st.selectbox('Module to update',module_list)
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
        st.write('Proceed to update')
        faculty = st.selectbox('Faculty',df_all['Name'])
        if faculty == df_Schedule[(df_Schedule['Location']==location) & (df_Schedule['Batch']==batch) & (df_Schedule['Course']==module)]['Faculty'].iloc[0]:
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
                print(ac)
                #write_to_sheet('Schedule',list(['DSE-FT',location,str(start_date), str(end_date),batch,module,faculty,ac,'9:30:00 AM', '5:00:00 PM','Residency',7,7,Remarks]),SPREADSHEET_ID) 
                update_batchdata(location,batch,start_date,end_date,module,faculty,Remarks)
                dates=pd.date_range(start= start_date, end=end_date)
                for d in dates:
                    write_to_sheet('Calendar',list([faculty,str(d),'9:30:00 AM', '5:00:00 PM',module,location,batch]),SPREADSHEET_ID)
                set_batch(batch,location,str(start_date),module)

    elif btype == 'Online':
        batches = all_onlineactive()
        batch = st.selectbox('Batch',list(batches.Batch))
        df = read_online_schedule(batch)
        st.write(df)
        module_list = list(df_Modules['Module Name'])
        print(module_list)
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
                            write_to_sheet('Calendar',list([faculty,str(d),'8:00:00 PM', '10:00:00 PM',module,'Online',batch]),SPREADSHEET_ID)
                        else:
                            write_to_sheet('Calendar',list([faculty,str(d),'9:30:00 AM', '1:30:00 PM',module,'Online',batch]),SPREADSHEET_ID)
                    set_onlinebatch(batch,str(start_date),module)
