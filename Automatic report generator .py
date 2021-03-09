from __future__ import print_function
from datetime import datetime, timedelta
from typing import OrderedDict
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import re 
import csv
import pickle
import gspread
from collections import OrderedDict 
import gspread_formatting as gsf


class Date:
    #this class will help to square away dates 
    def __init__(self):
        self.today = datetime.today()
        self.yesterday = datetime.today() - timedelta(days=1)
        self.last_week = datetime.today() - timedelta(days=7)
    def time_dif_by_day(self,days_ago):
        delta_day = self.today - timedelta(days=days_ago)
        return(delta_day.date())
    def day_after(self,date,num_days):
        delta_day = date + timedelta(days=num_days)
        return(delta_day)
    def day_before(self,date,num_days):
        delta_day = date - timedelta(days=num_days)
        return(delta_day)
    def nearest_monday(self):
        #days_from = {}
        n = self.today.weekday()
        if n < 5:
            n += 7
        #potential dict of the days and their values
        '''days_from['monday'] = n 
        days_from['tuesday'] = n - 1
        days_from['wednesday'] = n - 2 
        days_from['thursday'] = n - 3 
        days_from['friday'] = n - 4
        days_from['saturday'] = n - 5'''
        return self.time_dif_by_day(n)
    def string_to_dateobj(self, entry):
        dt_object = datetime.strptime(entry,'%Y-%m-%d')
        return dt_object
    def convert_time(self,string):
        string = re.sub(r'\..*$','',string )
        dt_object = datetime.strptime(string,'%Y-%m-%dT%H:%M:%S')
        dt_object = dt_object.strftime('%m/%d/%Y %I:%M %p')
        return dt_object
    def secs_to_hrs_mins_secs(self, value):
        hours = value // 3600
        minutes = (value - (hours*3600)) // 60
        seconds = value - (hours *3600 + minutes*60)
        return '{} hr {} mins {} secs'.format(round(hours),round(minutes), round(seconds))

def google_authenticate():
    os.chdir('d:\\code library\\reports\\') 
    #this is how we authenticate to the api and sdk 
    SCOPES = ['https://www.googleapis.com/auth/admin.reports.audit.readonly', 
    'https://www.googleapis.com/auth/admin.reports.usage.readonly',
    'https://www.googleapis.com/auth/admin.directory.user.readonly']
    creds = None
    #   The file token.pickle stores the user's access and refresh tokens, and is
    #   created automatically when the authorization flow completes for the first
    #   time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
            return creds
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            return creds
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
            return creds

def create_class_list(analyzed_data):
    # creates a class list with teachers and their students based on data given from week report or report_raw
    # not used with the main application but a useful tool that can quickly create class lists if asked
    dict_of_teach = {}
    for user, data in analyzed_data.items():
        if '@' in user:
            if data['teacher'] not in dict_of_teach:
                dict_of_teach[data['teacher']] = []
            dict_of_teach[data['teacher']].append(user)
    dict1 =  OrderedDict(sorted(dict_of_teach.items()))
    for teacher, students in dict1.items():
        print(teacher + ': ')
        for student in students:
            print('     ' + student)

def dict_format_creator():
    '''this is an in house solution to create the custom summary page that is used in the report it entirely reliese on being programmed correctly and not taking arguments'''

    # this creates a dictionary that formats the summary page. Primarily it tells the batch update what needs to be merged and how the merged cells needs to look. If they need to have special borders
    # the dictionary's keys are just labels to tell you what the format is trying to target the values are what the format will be. 
    # the first 4 values are start and end values for columns (read the devlopment page for more info regarding specifics )
    # the next value (5) indicates what style they want the border to be
    # the last value is a list of top,bottom,left,and right. It tells the program if a specific side needs to be more bolded than the rest. If left blank it defaults to the style choice in value 5
    # this is what allows the summary page to have distinct sections in regards to borders, some format requests using gspread were also made for the sake of consitency so you may find redudancy in the 
    # google form generator section

    #this first dictionary intializes with values for the headers larger than a collumn or 2 rows. They are labeled here for the sake of the reader and not for the program
    #these have to be put in by hand because they are unique to themselves and do not repeat
    Dict_for_format = {
    'top': [0,4,0,15,'THICK',[]],
    'total number of students' : [4,7,4,6, 'MEDIUM',['left','right']],
    'activity summary for classroom' : [4,7,6,15,'MEDIUM',['left','right']],
    'alert list' : [7,46,0,4,'',[]],
    'Google Docs Data' : [25,28,4,8,'THICK',['top']],
    'usage summary of logins' : [25,28,8,15,'THICK',['top']],
    'Meet Report' : [46,49,5,15,'THICK',['top','left','right']]
    }

    # once the section headers are put together we need to do all the sections inbetween which by hand would need to be about 85 entries. To solve this problem 
    # I have used the inherent simlarity of certain sections to automate this process which basically means we look at the start and end of sections and use that to merge all the cells in a pattern 
    # consitent with the vision of the summary page. Look to devolpment page for more on this vision
    # this next section intializes all of the start and end values for x which is the row count and y which is the column count
    # there are (as of 3-1-21) 3 sections that are made using this process and thus 3 distinct x and y values need to be used however Y values overlap for the last 2 sections and it means only 1 y value is needed
    # the start value for rows or columns needs to start 1 row or column before the needed row (look at deoployment page for more detail)
    start_value_y = 4
    end_value_y_1 = 14
    end_value_y = 15
    start_value_y_3 = 5
   

    start_value_x_1 = 7
    end_value_x_1 = 25
    start_value_x_2 = 28
    end_value_x_2 = 46
    start_value_x_3 = 49
    end_value_x_3 = 67
    x = start_value_x_1
    y = start_value_y
    #this loop covers the total numbers of student and Classroom Summary sections E8:N25
    # the basic idea for all of the loops is to create entries in the dictionary that correspond the upper and outward bounds provided. 
    #This creates a uniform section in the areas provided with while 'Y =' loops indictating specific formatting needing to happen to values the fall in a specific column
    # the X value needs to to add 1 to Y everytime it hits its upper limit and the Y value needs to stop everything once its upper limit is met. This program basically creates the formatting column by column'
    # the default method merges together 2 rows in each collumn creating more disctinct boxes that are easier to read. You need to program which column needs specific formatting to create sections
    while y < 15:
        if x == end_value_x_1:
            x = start_value_x_1
            y+= 1
        if y == end_value_y_1:
            x = start_value_x_2
            y= start_value_y
            break
        while x < end_value_x_1 - 1:
            if y == end_value_y_1:
                x = 100
                break
            while y == start_value_y:
                Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['left']]
                x += 2
                if x == end_value_x_1:
                    y+= 1
                    x = start_value_x_1
            while y == 5:
                Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['right']]
                x += 2
                if x == end_value_x_1:
                    y+= 1
                    x = start_value_x_1
            while y == 6:
                Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['left']]
                x += 2
                if x == end_value_x_1:
                    y+= 1
                    x = start_value_x_1
            while y == 13:
                    Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['right']]
                    x += 2
                    if x == end_value_x_1:
                        y+= 1
                        x = start_value_x_1
            else:
                if y == end_value_y:
                    break
                Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',[]]
                x += 2
    
    
    
    #this loop covers google docs and login sections E29:O46
    x = start_value_x_2
    y = start_value_y
    while y < 15:
        if y == end_value_y:
            break
        if x == end_value_x_2:
                x = start_value_x_2
                y+= 1
        while x < end_value_x_2 - 1:
                if y == end_value_y:
                    break
                if x == 44:
                    Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['bottom']]
                    x = start_value_x_2
                    y += 1
                
                while y == 4:
                    if x == 44:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['bottom','left']]
                    else:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['left']]
                    x += 2
                    if x == end_value_x_2:
                        y+= 1
                        x = start_value_x_2

                while y == 8:
                    if x == 44:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['bottom','left']]
                    Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['left']]
                    x += 2
                    if x == end_value_x_2:
                        y+= 1
                        x = start_value_x_2
                while y == 14:
                    if x == 44:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['bottom','right']]
                    else:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['right']]
                    x += 2
                    if x == end_value_x_2:
                        y+= 1
                        x = start_value_x_2
                else:
                    if y == end_value_y:
                        break
                    if x == 44:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['bottom']]
                        x = start_value_x_2
                        y += 1
                    else:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',[]]
                        x += 2
    
    #this section covers meet report F50:O67
    x = start_value_x_3
    y= start_value_y_3
    while y < 15:
        if y == end_value_y:
            break
        if x == end_value_x_3:
                x = start_value_x_3
                y+= 1
        while x < end_value_x_3 - 1:
                if y == end_value_y:
                    break
                if x == 65:
                    Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['bottom']]
                    x = start_value_x_3
                    y += 1
                
                while y == 5:
                    if x == 65:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['bottom','left']]
                    else:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['left']]
                    x += 2
                    if x == end_value_x_3:
                        y+= 1
                        x = start_value_x_3
                while y == 14:
                    if x == 65:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['bottom','right']]
                    else:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['right']]
                    x += 2
                    if x == end_value_x_3:
                        y+= 1
                        x = start_value_x_3
                else:
                    if y == end_value_y:
                        break
                    if x == 65:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',['bottom']]
                        x = start_value_x_3
                        y += 1
                    else:
                        Dict_for_format['body_for Section' + str(x) +str(y)] = [x, x+2, y, y+1,'',[]]
                        x += 2
    #this eventually returns a dictionary with primed formatting data                
    return(Dict_for_format)

def config_settings(location,d):
    #this section of code takes information from csv file and turns it into a dictionary with the key being a single word and its value being a list of values.
    # d is used to tell if the function needs to take data from a user csv report. look at user generator for more info on this process   
    with open(location, mode='r') as f:
        reader = csv.DictReader(f)
        config = {}
        list = []
        #this allows me to pull just the top row's values and set them to fieldnames
        fieldnames = reader.fieldnames
        #this puts the key into the field and creates a list as a result
        for fieldname in fieldnames:
            config[fieldname] = []
        #this section takes data from each row and matches it's key and then adds it to the list 
        if d != 'users':
            for row in reader:
                for fieldname in fieldnames:
                    config[fieldname].append(row[fieldname])
        else:
            for row in reader:
                config[row['email']] = {'name': row['name'], 'teacher': row['teacher'], 'grade' : row['grade'], 'suspended': row['suspended']}
        #now returns a dictionary with keys as fieldnames and a list of values associated to those keys
        return [config,fieldnames]

def user_list_generator():
    ''' This function looks at all users on raleighoakcharter.org domain
        takes information regarding their OU and then places them into a local csv file''' 
    def data_writer(user,first):
        list = [] 
        # /.?([a-zA-Z0-9]+)/([a-zA-Z0-9\-_ \.]+$)
        for user in users.get('users'):
            #use regular expression to get just the the teachers name on the OU
            # if the OU structure changes this will need to be modified. The current strucutre as of (03-01-21) is \Students\grade\teacher
            # if that changes it could break something
            Class = re.search(r'/\.?([a-zA-Z0-9]+)/(\w+)$',user['orgUnitPath'])
            if Class == None:
                continue
            Class_ = Class.group(2)
            Grade = Class.group(1)
            list.append({'email':user['primaryEmail'],
            'teacher' : Class_,
            'grade' : Grade,
            'name': user['name']['fullName'],
            'suspended': user['suspended']})
        if first == True:
            f = open('users.csv','w')
            writer = csv.DictWriter(f, ['email','teacher','grade','name','suspended'])
            writer.writeheader()  
            f.close()        
        
        with open('users.csv', 'a') as f: 
            writer = csv.DictWriter(f, ['email','teacher','grade','name','suspended'])
            for l in list:
                writer.writerow(l)
 #this is the credientals that allow the admin sdk to verify who I am
    creds = google_authenticate()
    #this changes the cwd to the data section so it can easily pull data from those lists 
    os.chdir('D:\\data\\reports\\')
    #this creates the service where we can call the admin sdk for our reports
    service = build('admin', 'directory_v1', credentials=creds)
    #this calls the first report which is limited to 200 users by default. 
    users = service.users().list(customer='my_customer', viewType='admin_view').execute()
    #this takes the data into data writer where it puts the first 200 users into along with the header  
    data_writer(users, True)
    #if the json file reports a 'next page token' it will use it to continue past 200 users
    test = len(users.get('nextPageToken',''))

    # this loop completes only when the 
    while test > 1: 
        token = users.get('nextPageToken')
        last_token = users.get('nextPageToken','')
        users = service.users().list(customer='my_customer', viewType='admin_view', pageToken = token ).execute()
        data_writer(users,False)
        token = users.get('nextPageToken','')
        test = len(token)
        if last_token == token: 
            break

def CSV_Download_name(name, new_file, field):
    #not used in any part of this program but was thought to be useful somewhere down the line and is best kept here
    '''this function is meant to take data from a google sheet and make it directly into a csv file''' 
    # use creds to create a client to interact with the Google Drive API
    gc = gspread.service_account(filename='D:/code library/onboarding/Service_account_cred.json')
    sh = gc.open(name)

    # Find a workbook by name or position and open the first sheet
    worksheet = sh.sheet1

    #the fields will help to sort the csv and not take irrelavent data 
    fields = field

    #create new csv file and create header line with fields
    with open(new_file,'w') as file: 
        data = worksheet.get_all_records()
        writer = csv.DictWriter(file,fields,extrasaction='ignore')
    #create header
        writer.writeheader()
    #add rows until data stops
        writer.writerows(data)

def report_raw(date_entry):
    #this is the core of the week_report generator
    # this takes information from a user usage report filters it based off of filednames provided in data.csv and then exports a dicitonary with keys being 
    # since reports are limited to 200 results default with a maximum of 500 there needs to be built in logic to deal with the next page token request
    # this is why there is a first report which then provides 200 results and an nextpage token which is then used to generate the next 200 tokens
    # this code is basically able to handle 1 results to millions

    # this code adds every user to the data_log dictionary and then adds the relevant data
    def data_log_builder(results,data_log):
        for result in results.get('usageReports'):
            email = result['entity']['userEmail']
            activities = result['parameters']
            data_log[email] = {}
            
            for activity in activities:
                for fieldname in fieldnames: 
                    if activity['name'] == fieldname:
                        for key,value in activity.items():
                            if key == 'name':
                                x = value
                            data_log[email][x] = value
        return data_log
    
    
    date = Date()
    creds = google_authenticate()
    service = build('admin', 'reports_v1', credentials=creds)
    user_log = {}
    fieldnames = config_settings('d:/data/reports/data.csv','')[0]['Fieldnames']
    print('report is generating')


    #this calls the API to generate a user report with specific variables
    results = service.userUsageReport().get(userKey='all',date=date_entry).execute()
    if 'usageReports' not in results:
            return 'data is not ready for export'
    data_log ={}

    data_log = data_log_builder(results,data_log)
    while results.get('nextPageToken','') != '':
        token = results.get('nextPageToken','')
        results = service.activities().list(userKey='all',date=date_entry, pageToken = token).execute()
        data_log = data_log_builder(results,data_log)
    
    dd = date.string_to_dateobj(date_entry)
    data_log['date'] = {'date' : dd, 'weekday' : date.today.weekday()}
    return data_log

def activities_report_meet():
    # this is very similar in structure to report raw but pulls an activity report rather than an usage report
    # it also deals with next page tokens and should not need to be modified regardless of amount of users
    # unlike the usage report where I want several fields from each entry I only need a few and thus is specified in the code itself.
    # this report looks at every single time someone logs out of a google meet. It will tell youhow long they were in that meet. This allows the report to tell if people are active on a specific day
    def data_log_builder(results,data_log):
        for result in results.get('items'):
            for user in users:
                if '@' in user:
                    if result['actor'].get('email','') == user:
                        if data_log.get(user,'') == '':
                            data_log[user] = {}
                            data_log[user]['dates_of_meets'] = []
                        for events in result['events']:
                            for event in events['parameters']:
                                if event['name'] == 'duration_seconds':
                                    if data_log[user].get('time_in_meet',0) == 0:
                                        data_log[user]['time_in_meet'] = 0
                                    data_log[user]['time_in_meet'] += int(event['intValue'])
                                    data_log[user]['dates_of_meets'].append([date.convert_time(result['id']['time']) , int(event['intValue'])])
        return(data_log)

    date = Date()
    monday = date.nearest_monday()
    creds = google_authenticate()
    monday_ISO = str(monday) + 'T00:00:00.000Z'
    friday_ISO = str(date.day_after(monday,5)) + 'T00:00:00.000Z'
    print('activities report is generating')

    service = build('admin', 'reports_v1', credentials=creds)

    results = service.activities().list(userKey='all' ,applicationName = 'meet', startTime = monday_ISO, endTime = friday_ISO, eventName = 'call_ended').execute()
    users = config_settings('d:\\data\\reports\\users.csv','users')[0]
    data_log = {}

    data_log = data_log_builder(results,data_log)
    while results.get('nextPageToken','') != '':
        token = results.get('nextPageToken','')
        results = service.activities().list(userKey='all' ,applicationName = 'meet', startTime = monday_ISO, endTime = friday_ISO,eventName = 'call_ended', pageToken = token).execute()
        data_log = data_log_builder(results,data_log)
    return data_log
                    
def data_keys(analyzed_data):
    wr = analyzed_data
    date= Date()
    monday = date.nearest_monday()
    friday = str(date.day_after(monday,4))
    with open('D:\\data\\reports\\data_keys\\data_keys_values{}_{}.csv'.format(monday,friday),'w') as f:
        writer = csv.writer(f)
        for key, value in wr['data'].items():
            writer = csv.writer(f)
            writer.writerow([key,value])
    
def week_report():
    def report_to_report(report_1,report_2,day_of_week):
        date = Date()
        #report 1 is the first report used to create the analyzed data dict and thus needs to chronologically come first. 
        #report 2 will be analyzed against report 1 for changes in interaction time etc. 
        '''this is the intial report. Needs to be started first as it creates all the values needed for anaylsis'''
        if report_1 == 'data is not ready for export':
            return 'week is not ready'
        if report_2 == 'data is not ready for export':
            return 'week is not ready'
        #intializes the dictionary to hold all the data
        analyzed_data = {}
        #this generates a list of users from a csv file into a dictionary where the first key is their email and the values are: name, grade, teacher, suspend status 
        #you need to look at users in your data folder to see what terms have been used because they may vary. 
        users = config_settings('d:\\data\\reports\\users.csv','users')[0]
        #takes data from first report and generates the analyzed_data dictionary with base values 
        for user,data in report_1.items():
            if data.get('classroom:role','') == 'teacher':
                    continue   
            for email,values in users.items():
                if user == email:
                    if values['suspended'] == 'True':
                        continue
                    #this is the base analazed data dictionary
                    analyzed_data[user] = {'name':'none','email': user, 'classroom_log':
                {'monday':'none', 'mon_login' : 'none', 'tues_login' : 'none','tuesday':'none','wednes_login' : 'none','wednesday':'none','thurs_login' : 'none',
                'thursday':'none','fri_login' : 'none','friday': 'none','missed_days': 0, 'active_days':0},
                'teacher':'none', 'grade': 'none', 'drive_log': {'mon_created': 'none', 'tues_created': 'none', 
                'wednes_created': 'none','thurs_created': 'none','fri_created':'none'}, 'login_log' : {'monday':'none', 'mon_login' : 'none', 'tues_login' : 'none','tuesday':'none','wednes_login' : 'none','wednesday':'none','thurs_login' : 'none',
                'thursday':'none','fri_login' : 'none','friday': 'none','missed_days': 0, 'active_days':0}}
                    for key, value in values.items():
                        analyzed_data[user][key] = value
                else:
                    continue
            # this now starts the filling of the data 
            if user in analyzed_data:
                for user_monday,data_monday in report_2.items():
                    if user_monday == user:
                        #analyzing google classroom interaction data 
                        if data_monday['classroom:last_interaction_time'] == data['classroom:last_interaction_time']:
                            analyzed_data[user]['classroom_log'][day_of_week + '_login'] = date.convert_time(data_monday['classroom:last_interaction_time'])
                            analyzed_data[user]['classroom_log'][day_of_week + 'day'] = 'No'
                            analyzed_data[user]['classroom_log']['missed_days'] += 1
                        else:
                            analyzed_data[user]['classroom_log'][day_of_week + '_login'] = date.convert_time(data_monday['classroom:last_interaction_time'])
                            if re.search(r'(\d\d\d\d-\d\d-\d\d)T', data_monday['classroom:last_interaction_time']).group(1) == str(report_2['date']['date'].date()):
                                analyzed_data[user]['classroom_log'][day_of_week + 'day'] = 'Yes'
                                analyzed_data[user]['classroom_log']['active_days'] += 1
                            else:
                                date_one_day = report_2['date']['date'] + timedelta(days=1)
                                if re.search(r'(\d\d\d\d-\d\d-\d\d)T', data_monday['classroom:last_interaction_time']).group(1) == str(date_one_day.date()):
                                    analyzed_data[user]['classroom_log'][day_of_week + 'day'] = 'Yes'
                                    analyzed_data[user]['classroom_log']['active_days'] += 1
                                else:
                                    analyzed_data[user]['classroom_log'][day_of_week + 'day'] = 'No'
                                    analyzed_data[user]['classroom_log']['missed_days'] += 1
                        
                        # begin login data analysis
                        if data_monday.get('accounts:last_login_time','') == data.get('accounts:last_login_time',''):
                            analyzed_data[user]['login_log'][day_of_week + '_login'] = date.convert_time(data_monday.get('accounts:last_login_time',''))
                            analyzed_data[user]['login_log'][day_of_week + 'day'] = 'No'
                            analyzed_data[user]['login_log']['missed_days'] += 1
                        else:
                            analyzed_data[user]['login_log'][day_of_week + '_login'] = date.convert_time(data_monday['accounts:last_login_time'])
                            if re.search(r'(\d\d\d\d-\d\d-\d\d)T', data_monday['accounts:last_login_time']).group(1) == str(report_2['date']['date'].date()):
                                analyzed_data[user]['login_log'][day_of_week + 'day'] = 'Yes'
                                analyzed_data[user]['login_log']['active_days'] += 1
                                continue
                            else:
                                date_one_day = report_2['date']['date'] + timedelta(days=1)
                                if re.search(r'(\d\d\d\d-\d\d-\d\d)T', data_monday['accounts:last_login_time']).group(1) == str(date_one_day.date()):
                                    analyzed_data[user]['login_log'][day_of_week + 'day'] = 'Yes'
                                    analyzed_data[user]['login_log']['active_days'] += 1
                                else:
                                    analyzed_data[user]['login_log'][day_of_week + 'day'] = 'No'
                                    analyzed_data[user]['login_log']['missed_days'] += 1

                    analyzed_data[user]['drive_log'][day_of_week + '_created'] = int(data_monday.get('drive:num_google_documents_created',0)) + int(data_monday.get('drive:num_google_forms_created',0)) + int(data_monday.get('drive:num_google_presentations_created',0)) + int(data_monday.get('drive:num_google_spreadsheets_created',0))
                    analyzed_data[user]['drive_log'][day_of_week + '_viewed'] = int(data_monday.get('drive:num_google_documents_viewed',0)) + int(data_monday.get('drive:num_google_forms_viewed',0)) + int(data_monday.get('drive:num_google_presentations_viewed',0)) + int(data_monday.get('drive:num_google_spreadsheets_viewed',0))
                    analyzed_data[user]['drive_log'][day_of_week + '_edited'] = int(data_monday.get('drive:num_google_documents_edited',0)) + int(data_monday.get('drive:num_google_forms_edited',0)) + int(data_monday.get('drive:num_google_presentations_edited',0)) + int(data_monday.get('drive:num_google_spreadsheets_edited',0))
                    analyzed_data[user]['drive_log']['total docs created'] = 0
                    analyzed_data[user]['drive_log']['total docs viewed'] = 0
                    analyzed_data[user]['drive_log']['total docs edited'] = 0
                    if analyzed_data[user]['drive_log'][day_of_week + '_created'] != 'none':
                        analyzed_data[user]['drive_log']['total docs created'] += analyzed_data[user]['drive_log'][day_of_week + '_created']
                    if analyzed_data[user]['drive_log'][day_of_week + '_viewed'] != 'none':
                        analyzed_data[user]['drive_log']['total docs viewed'] += analyzed_data[user]['drive_log'][day_of_week + '_viewed']
                    if analyzed_data[user]['drive_log'][day_of_week + '_edited'] != 'none':
                        analyzed_data[user]['drive_log']['total docs edited'] += analyzed_data[user]['drive_log'][day_of_week + '_edited']
 
        return analyzed_data
    
    def report_to_data(report_1, analyzed_data,day_of_week,previous_day_of_week):
        date = Date()
        if report_1 == 'data is not ready for export':
            return 'week is not ready'
        for user,data in analyzed_data.items():
            if user in report_1:
                for user_anyday,data_anyday in report_1.items():
                    if user_anyday == user:
                        #begin classroom log data anaylsis
                        if date.convert_time(data_anyday['classroom:last_interaction_time']) == data['classroom_log'][previous_day_of_week + '_login']:
                            data['classroom_log'][day_of_week + '_login'] = data['classroom_log'][previous_day_of_week + '_login']
                            data['classroom_log'][day_of_week + 'day'] = 'No'
                            data['classroom_log']['missed_days'] += 1
                        else:
                            data['classroom_log'][day_of_week + '_login'] = date.convert_time(data_anyday['classroom:last_interaction_time'])
                            if re.search(r'(\d\d\d\d-\d\d-\d\d)T', data_anyday['classroom:last_interaction_time']).group(1) == str(report_1['date']['date'].date()):
                                data['classroom_log']['active_days'] += 1
                                data['classroom_log'][day_of_week + 'day'] = 'Yes'
                            else:
                                date_one_day = report_1['date']['date'] + timedelta(days=1)
                                if re.search(r'(\d\d\d\d-\d\d-\d\d)T', data_anyday['classroom:last_interaction_time']).group(1) == str(date_one_day.date()):
                                    data['classroom_log'][day_of_week + 'day'] = 'Yes'
                                    data['classroom_log']['active_days'] += 1
                                else:
                                    data['classroom_log'][day_of_week + 'day'] = 'No'
                                    data['classroom_log']['missed_days'] += 1
                        #begin login data anaylsis
                        if date.convert_time(data_anyday.get('accounts:last_login_time','')) == data['login_log'][previous_day_of_week + '_login']:
                            data['login_log'][day_of_week + '_login'] = data['login_log'][previous_day_of_week + '_login']
                            data['login_log'][day_of_week + 'day'] = 'No'
                            data['login_log']['missed_days'] += 1
                        else:
                            data['login_log'][day_of_week + '_login'] = date.convert_time(data_anyday['accounts:last_login_time'])
                            if re.search(r'(\d\d\d\d-\d\d-\d\d)T', data_anyday['accounts:last_login_time']).group(1) == str(report_1['date']['date'].date()):
                                data['login_log'][day_of_week + 'day'] = 'Yes'
                                data['login_log']['active_days'] += 1
                                continue
                            else:
                                date_one_day = report_1['date']['date'] + timedelta(days=1)
                                if re.search(r'(\d\d\d\d-\d\d-\d\d)T', data_anyday['accounts:last_login_time']).group(1) == str(date_one_day.date()):
                                    data['login_log'][day_of_week + 'day'] = 'Yes'
                                    data['login_log']['active_days'] += 1
                                else:
                                    data['login_log'][day_of_week + 'day'] = 'No'
                                    data['login_log']['missed_days'] += 1

                        analyzed_data[user]['drive_log'][day_of_week + '_created'] = int(data_anyday.get('drive:num_google_documents_created',0)) + int(data_anyday.get('drive:num_google_forms_created',0)) + int(data_anyday.get('drive:num_google_presentations_created',0)) + int(data_anyday.get('drive:num_google_spreadsheets_created',0))
                        analyzed_data[user]['drive_log'][day_of_week + '_viewed'] = int(data_anyday.get('drive:num_google_documents_viewed',0)) + int(data_anyday.get('drive:num_google_forms_viewed',0)) + int(data_anyday.get('drive:num_google_presentations_viewed',0)) + int(data_anyday.get('drive:num_google_spreadsheets_viewed',0))
                        analyzed_data[user]['drive_log'][day_of_week + '_edited'] = int(data_anyday.get('drive:num_google_documents_edited',0)) + int(data_anyday.get('drive:num_google_forms_edited',0)) + int(data_anyday.get('drive:num_google_presentations_edited',0)) + int(data_anyday.get('drive:num_google_spreadsheets_edited',0))
                        if analyzed_data[user]['drive_log'][day_of_week + '_created'] != 'none':
                            analyzed_data[user]['drive_log']['total docs created'] += analyzed_data[user]['drive_log'][day_of_week + '_created']
                        if analyzed_data[user]['drive_log'][day_of_week + '_viewed'] != 'none':
                             analyzed_data[user]['drive_log']['total docs viewed'] += analyzed_data[user]['drive_log'][day_of_week + '_viewed']
                        if analyzed_data[user]['drive_log'][day_of_week + '_edited'] != 'none':
                            analyzed_data[user]['drive_log']['total docs edited'] += analyzed_data[user]['drive_log'][day_of_week + '_edited']
        return analyzed_data
    
    def further_data_anaylasis(analyzed_data,date_entry):
        #this functions takes all the raw analyzed data adds the "data" section to it where it calculates averages and other pieces based off of group trends
        #data does not carry stats about specific users but overall trends
        print('analyzing all data')
        #intializing all lists and dictionaries used in calculations
        grade_list = []
        analyzed_data['data'] = {}
        analyzed_data['data']['# of students'] = 0
        analyzed_data['data']['total missed days in classroom'] = 0
        analyzed_data['data']['total active days in classroom'] = 0
        analyzed_data['data']['total docs created'] = 0
        analyzed_data['data']['total docs viewed'] = 0
        analyzed_data['data']['total docs edited'] = 0
        analyzed_data['data']['total missed days by login'] = 0
        analyzed_data['data']['total active days by login'] = 0
        analyzed_data['data']['total active days in meet'] = 0
        analyzed_data['data']['total missed days in meet'] = 0
        analyzed_data['data']['alert_list'] = []
        day_list = ['monday','tuesday','wednesday','thursday','friday']
        cs = config_settings('d:\\data\\reports\\data.csv','')
        #iterating over every user sorting them by grade and calcuating totals based on activity.
        for user,data in analyzed_data.items():
            date = Date()
            if '@' in user:

                for day in day_list:
                    if data['meet_report'].get(day,'') == '':
                        data['meet_report'][day]= 'No'

                #calculating out plan b students
                for person in cs[0]['mon_tues']:
                    data_list = ['classroom_log','login_log', 'meet_report']
                    if person +'@raleighoakcharter.org' == user:
                        data['plan_b'] ='monday & tuesday'
                        for log in data_list:
                            if analyzed_data[user][log]['monday'] == 'Yes':
                                analyzed_data[user][log]['monday'] = 'Yes'
                            else: 
                                analyzed_data[user][log]['active_days']+= 1
                                analyzed_data[user][log]['missed_days']-= 1
                                analyzed_data[user][log]['monday'] = 'Yes'
                            if analyzed_data[user][log]['tuesday'] == 'Yes':
                                analyzed_data[user][log]['tuesday'] = 'Yes'
                            else: 
                                analyzed_data[user][log]['tuesday'] = 'Yes'
                                analyzed_data[user][log]['active_days']+= 1
                                analyzed_data[user][log]['missed_days']-= 1
                for person in cs[0]['thurs_fri']:
                    if person +'@raleighoakcharter.org' == user:
                        data['plan_b']='thursday & friday'
                        for log in data_list:
                            if analyzed_data[user][log]['friday'] == 'Yes':
                                analyzed_data[user][log]['friday'] = 'Yes'
                            else: 
                                analyzed_data[user][log]['active_days']+= 1
                                analyzed_data[user][log]['missed_days']-= 1
                                analyzed_data[user][log]['friday'] = 'Yes'
                            if analyzed_data[user][log]['thursday'] == 'Yes':
                                analyzed_data[user][log]['thursday'] = 'Yes'
                            else: 
                                analyzed_data[user][log]['thursday'] = 'Yes'
                                analyzed_data[user][log]['active_days']+= 1
                                analyzed_data[user][log]['missed_days']-= 1
                if data.get('plan_b','') =='':
                    data['plan_b'] = 'No'
                grade = data['grade']
                #teacher = data['teacher']
                #grade is equivilanet to the OU of the user. Students is the root of the student drive where fake student accounts live so they are skipped
                if grade == 'Students':
                    continue
                #create total number of students in report
                analyzed_data['data']['# of students'] += 1
                #begin classroom analysis by school and then by sub divisions
                analyzed_data['data']['total missed days in classroom'] += data['classroom_log']['missed_days']
                analyzed_data['data']['total active days in classroom'] += data['classroom_log']['active_days']
                #grade based analsis like averages per grade, etc. Potential to do it by teacher but not asked to as of now
                if 'total of active days in classroom by ' + grade not in analyzed_data['data']:
                    analyzed_data['data']['total of active days in classroom by ' + grade] = 0
                    #this calcuates the total of students in the grade
                    analyzed_data['data']['total # of students in ' + grade] = 0
                    grade_list.append(grade)
                if 'total of missed days in classroom by ' + grade not in analyzed_data['data']:
                    analyzed_data['data']['total of missed days in classroom by ' + grade] = 0
                if analyzed_data['data'].get('# of non-active students in classroom in ' + grade,'') == '':
                    analyzed_data['data']['# of non-active students in classroom in ' + grade] = 0
                if analyzed_data[user]['classroom_log']['missed_days'] == 5:
                    analyzed_data['data']['# of non-active students in classroom in ' + grade] += 1
                analyzed_data['data']['total of active days in classroom by ' + grade] += data['classroom_log']['active_days']
                analyzed_data['data']['total of missed days in classroom by ' + grade] += data['classroom_log']['missed_days']
                analyzed_data['data']['total # of students in ' + grade] += 1
                
                #begin login data anylasis
                analyzed_data['data']['total missed days by login'] += data['login_log']['missed_days']
                analyzed_data['data']['total active days by login'] += data['login_log']['active_days']
                
                #grade based analsis like averages per grade, etc. Potential to do it by teacher but not asked to as of now
                if 'total of active days by login of ' + grade not in analyzed_data['data']:
                    analyzed_data['data']['total of active days by login of ' + grade] = 0
                if 'total of missed days by login of ' + grade not in analyzed_data['data']:
                    analyzed_data['data']['total of missed days by login of ' + grade] = 0
                if analyzed_data['data'].get('# of non-active students from login ' + grade,'') == '':
                        analyzed_data['data']['# of non-active students from login ' + grade] = 0
                if analyzed_data[user]['login_log']['missed_days'] == 5:
                    analyzed_data['data']['# of non-active students from login ' + grade] += 1
                analyzed_data['data']['total of active days by login of ' + grade] += data['login_log']['active_days']
                analyzed_data['data']['total of missed days by login of ' + grade] += data['login_log']['missed_days']
                if analyzed_data[user]['login_log']['missed_days'] == 5:
                    analyzed_data[user]['login_log']['alert'] = 'High'
                if analyzed_data[user]['login_log']['missed_days'] <= 4:
                    analyzed_data[user]['login_log']['alert'] = 'Moderate'
                if analyzed_data[user]['login_log']['missed_days'] <= 2:
                    analyzed_data[user]['login_log']['alert'] = 'Low'



                if 'total docs edited by ' + grade not in analyzed_data['data']:
                    analyzed_data['data']['total docs edited by ' + grade] =0
                if 'total docs viewed by ' + grade not in analyzed_data['data']:
                    analyzed_data['data']['total docs viewed by ' + grade]=0
                if 'total docs created by ' + grade not in analyzed_data['data']:
                    analyzed_data['data']['total docs created by ' + grade]=0
                analyzed_data['data']['total docs created by ' + grade] += data['drive_log']['total docs created']
                analyzed_data['data']['total docs viewed by ' + grade] += data['drive_log']['total docs viewed']
                analyzed_data['data']['total docs edited by ' + grade] += data['drive_log']['total docs edited']



                #begin meet anlaysis
                if data['meet_report'].get('missed_days','') == '':
                     data['meet_report']['missed_days'] = 0
                if data['meet_report'].get('active_days','') == '':  
                       data['meet_report']['active_days'] = 0
                if analyzed_data['data'].get('# of non-active students in meet','') == '':
                    analyzed_data['data']['# of non-active students in meet'] = 0
                if analyzed_data[user]['meet_report']['missed_days'] == 5:
                    analyzed_data['data']['# of non-active students in meet'] += 1
                analyzed_data['data']['total missed days in meet'] += data['meet_report']['missed_days']
                analyzed_data['data']['total active days in meet'] += data['meet_report']['active_days']

                if 'total of active days in meet by ' + grade not in analyzed_data['data']:
                    analyzed_data['data']['total of active days in meet by ' + grade] = 0
                if 'total of missed days in meet by ' + grade not in analyzed_data['data']:
                    analyzed_data['data']['total of missed days in meet by ' + grade] = 0
                if analyzed_data['data'].get('# of non-active students in meet in ' + grade,'') == '':
                    analyzed_data['data']['# of non-active students in meet in ' + grade] = 0
                if analyzed_data[user]['meet_report']['missed_days'] == 5:
                    analyzed_data['data']['# of non-active students in meet in ' + grade] += 1
                if analyzed_data['data'].get('raw_time_in_meet_' + grade,'') =='':
                     analyzed_data['data']['raw_time_in_meet_' + grade] = 0
                for day in day_list:
                    if data['meet_report'].get(day + ' time','') == '':
                        continue
                    else:
                        analyzed_data['data']['raw_time_in_meet_'+grade] += data['meet_report'][day + ' time']

                if analyzed_data[user]['meet_report']['missed_days'] <= 4:
                    analyzed_data[user]['meet_report']['alert'] = 'Moderate'
                    analyzed_data[user]['meet_report']['alert_reason'] = 'Missed {} days'.format(analyzed_data[user]['meet_report']['missed_days'])
                if analyzed_data[user]['meet_report']['missed_days'] <= 2:
                    analyzed_data[user]['meet_report']['alert'] = 'Low'
                    analyzed_data[user]['meet_report']['alert_reason'] = 'Missed {} days'.format(analyzed_data[user]['meet_report']['missed_days'])
                analyzed_data['data']['total of active days in meet by ' + grade] += data['meet_report']['active_days']
                analyzed_data['data']['total of missed days in meet by ' + grade] += data['meet_report']['missed_days']

                #this adds together all documents created, viewed, and edited
                for key,value in analyzed_data[user]['drive_log'].items():
                    if value =='none':
                        continue
                    if '_created' in key:
                        analyzed_data['data']['total docs created'] += value
                    if '_viewed' in key:
                        analyzed_data['data']['total docs viewed'] += value
                    if '_edited' in key:
                        analyzed_data['data']['total docs edited'] += value
                
                #this analysis the date at the last time of login and adds it to each user's data profile
                Last_class_login = date_entry - datetime.strptime(analyzed_data[user]['classroom_log']['fri_login'], '%m/%d/%Y %I:%M %p').date()
                if int(re.search(r'(\d+)',str(Last_class_login)).group(1)) > 30:
                    analyzed_data[user]['days_since_last_classroom_interaction'] = '30+'
                else:
                    analyzed_data[user]['days_since_last_classroom_interaction'] = re.search(r'(\d+)',str(Last_class_login)).group(1)


                #start of alert generator this will grade the students activity and set an alert level for the student
                #This also calculates the # of active students in the week.
                #critera will be based off of a 3 grade scale High, moderate, low. High alert will mean no interaction for 4 days
                #moderate will mean no interaction for 2-3 days and low means 1 or less days missed. Other criteria can be added but for now this is all
                if analyzed_data['data'].get('# of non-active students in classroom','') == '':
                        analyzed_data['data']['# of non-active students in classroom'] = 0
                if analyzed_data[user]['classroom_log']['missed_days'] == 5:
                    analyzed_data['data']['# of non-active students in classroom'] += 1
                    analyzed_data[user]['alert'] = 'No activity'
                if analyzed_data['data'].get('# of non-active students from login','') == '':
                        analyzed_data['data']['# of non-active students from login'] = 0
                if analyzed_data[user]['login_log']['missed_days'] == 5:
                    analyzed_data['data']['# of non-active students from login'] += 1
                if analyzed_data[user]['classroom_log']['missed_days'] > 3:
                    analyzed_data[user]['alert'] = 'High'
                    analyzed_data['data']['alert_list'].append(analyzed_data[user]['name'])
                    continue
                if analyzed_data[user]['classroom_log']['missed_days'] > 1:
                    analyzed_data[user]['alert'] = 'Moderate'
                    continue
                if analyzed_data[user]['classroom_log']['missed_days'] < 2:
                    analyzed_data[user]['alert'] = 'Low'
                    continue
                   
        #creating averages based off of data
        analyzed_data['data']['average active days in classroom'] = round(analyzed_data['data']['total active days in classroom'] / analyzed_data['data']['# of students'],2)
        analyzed_data['data']['average active days in meet'] = round(analyzed_data['data']['total active days in meet'] / analyzed_data['data']['# of students'],2)
        analyzed_data['data']['average active days by login'] = round(analyzed_data['data']['total active days by login'] / analyzed_data['data']['# of students'],2)
        analyzed_data['data']['average docs viewed'] = round(analyzed_data['data']['total docs viewed'] / analyzed_data['data']['# of students'],2)
        analyzed_data['data']['average docs edited'] = round(analyzed_data['data']['total docs edited'] / analyzed_data['data']['# of students'],2)
        analyzed_data['data']['average docs created'] = round(analyzed_data['data']['total docs created'] / analyzed_data['data']['# of students'],2)
        analyzed_data['data']['% of active students in classroom'] = 1 - (analyzed_data['data']['# of non-active students in classroom'] / analyzed_data['data']['# of students'])
        analyzed_data['data']['% of active students by login'] = 1 - (analyzed_data['data']['# of non-active students from login'] / analyzed_data['data']['# of students'])
        analyzed_data['data']['% of active students in meet'] = 1 - (analyzed_data['data']['# of non-active students in meet'] / analyzed_data['data']['# of students'])
        for grade in grade_list:
            
        
            #calculate averages for services
            analyzed_data['data']['average missed days in classroom by ' + grade] = round(analyzed_data['data']['total of missed days in classroom by ' + grade] / analyzed_data['data']['total # of students in ' + grade],2)
            analyzed_data['data']['average active days in classroom by ' + grade] = round(analyzed_data['data']['total of active days in classroom by ' + grade] / analyzed_data['data']['total # of students in ' + grade],2)
            analyzed_data['data']['average missed days by login of ' + grade] = round(analyzed_data['data']['total of missed days by login of ' + grade] / analyzed_data['data']['total # of students in ' + grade],2)
            analyzed_data['data']['average active days by login of ' + grade] = round(analyzed_data['data']['total of active days by login of ' + grade] / analyzed_data['data']['total # of students in ' + grade],2)
            analyzed_data['data']['average missed days in meet by ' + grade] = round(analyzed_data['data']['total of missed days in meet by ' + grade] / analyzed_data['data']['total # of students in ' + grade],2)
            analyzed_data['data']['average active days in meet by ' + grade] = round(analyzed_data['data']['total of active days in meet by ' + grade] / analyzed_data['data']['total # of students in ' + grade],2)
            
            #calculate deltas and percentage changes for classroom 
            percent_active = 1 - (analyzed_data['data']['# of non-active students in classroom in ' + grade] / analyzed_data['data']['total # of students in ' + grade])
            analyzed_data['data']['% of active students in classroom in ' + grade] = '{}%'.format(round(percent_active*100,2))
            percent = ((analyzed_data['data']['average active days in classroom by ' + grade] / analyzed_data['data']['average active days in classroom']) - 1)*100
            analyzed_data['data']['% delta from school average in classroom for ' +grade] = '{}%'.format(round(percent,2))
            percent = ((percent_active / analyzed_data['data']['% of active students in classroom']) - 1)*100
            analyzed_data['data']['% delta from school average of active students in classroom for ' +grade] = '{}%'.format(round(percent,2))

            #calculate deltas and percentage changes for logins

            percent_active = 1 - (analyzed_data['data']['# of non-active students from login ' + grade] / analyzed_data['data']['total # of students in ' + grade])
            analyzed_data['data']['% of active students by login in ' + grade] = '{}%'.format(round(percent_active*100,2))
            percent = ((analyzed_data['data']['average active days by login of ' + grade] / analyzed_data['data']['average active days by login']) - 1)*100
            analyzed_data['data']['% delta from school average of login for ' +grade] = '{}%'.format(round(percent,2))
            percent = ((percent_active / analyzed_data['data']['% of active students by login']) - 1)*100
            analyzed_data['data']['% delta from school average of active students from logins for ' +grade] = '{}%'.format(round(percent,2))

            #calculate deltas and percentage changes for meeet

            percent_active = 1 - (analyzed_data['data']['# of non-active students in meet in ' + grade] / analyzed_data['data']['total # of students in ' + grade])
            analyzed_data['data']['% of active students in meet in ' + grade] = '{}%'.format(round(percent_active*100,2))
            percent = ((analyzed_data['data']['average active days in meet by ' + grade] / analyzed_data['data']['average active days in meet']) - 1)*100
            analyzed_data['data']['% delta from school average in meet by ' +grade] = '{}%'.format(round(percent,2))
            percent = ((percent_active / analyzed_data['data']['% of active students in meet']) - 1)*100
            analyzed_data['data']['% delta from school average of active students in meet for ' +grade] = '{}%'.format(round(percent,2))
            analyzed_data['data']['total_time_in_meet_'+grade] = date.secs_to_hrs_mins_secs(analyzed_data['data']['raw_time_in_meet_' + grade])
            analyzed_data['data']['average_time_in_meet_' + grade] = date.secs_to_hrs_mins_secs((analyzed_data['data']['raw_time_in_meet_' + grade]/analyzed_data['data']['total # of students in ' + grade]))
            analyzed_data['data']['average_time_in_meet_raw_' + grade] = analyzed_data['data']['raw_time_in_meet_' + grade]/analyzed_data['data']['total # of students in ' + grade]

        percent = (1 - (analyzed_data['data']['# of non-active students in classroom'] / analyzed_data['data']['# of students']))*100
        analyzed_data['data']['% of active students in classroom'] = '{}%'.format(round(percent,2))
        percent = (1 - (analyzed_data['data']['# of non-active students from login'] / analyzed_data['data']['# of students']))*100
        analyzed_data['data']['% of active students by login'] = '{}%'.format(round(percent,2))
        percent = (1 - (analyzed_data['data']['# of non-active students in meet'] / analyzed_data['data']['# of students']))*100
        analyzed_data['data']['% of active students in meet'] = '{}%'.format(round(percent,2))
        return analyzed_data
    
    def meet_report(analyzed_data):

        date = Date()
        meet_report = activities_report_meet()
        
        monday =  date.nearest_monday()
        tuesday = date.day_after(monday,1)
        tuesday =  tuesday.strftime('%m/%d/%Y')
        wednesday = date.day_after(monday,2)
        wednesday =  wednesday.strftime('%m/%d/%Y')
        thursday = date.day_after(monday,3)
        thursday =  thursday.strftime('%m/%d/%Y')
        friday = date.day_after(monday,4)
        friday =  friday.strftime('%m/%d/%Y')
        monday = monday.strftime('%m/%d/%Y')

        date_list = {'monday' :monday,'tuesday' :tuesday,'wednesday' :wednesday,'thursday' :thursday,'friday' :friday}


        for user,non in analyzed_data.items():
            if '@' in user:
                for user_1,data in meet_report.items():
                    if user == user_1:
                        if analyzed_data[user].get('meet_report','') == '':
                            analyzed_data[user]['meet_report'] = {}
                            analyzed_data[user]['meet_report']['total_time_in_meets'] = date.secs_to_hrs_mins_secs(data['time_in_meet'])
                        for entry in data['dates_of_meets']:
                            for dates,value in date_list.items():
                                if re.search(r'([\d/]+)',entry[0]).group(1) == value:
                                    if analyzed_data[user]['meet_report'].get(dates + ' time','') == '': 
                                        analyzed_data[user]['meet_report'][dates + ' time'] = 0
                                    analyzed_data[user]['meet_report'][dates + ' time'] += entry[1]
                                else: 
                                    if analyzed_data[user]['meet_report'].get(dates + ' time','') == '':
                                        analyzed_data[user]['meet_report'][dates + ' time'] = 0
                                    continue
        for user,data in analyzed_data.items():
            if '@' in user:
                if data.get('meet_report', '')== '':
                    date = Date()
                    data['meet_report'] = {'alert':'High', 'alert_reason': 'no logins','missed_days':5, 'active_days':0 ,'total_time_in_meets':date.secs_to_hrs_mins_secs(0)}
                    for date in date_list:
                        if data['meet_report'].get(date + ' time', '') == '':
                            data['meet_report'][date + ' time'] = 0
                    continue
                for date in date_list.keys():
                    if data['meet_report'].get(date + ' time') == 0:
                        if data['meet_report'].get('missed_days','')=='':
                            data['meet_report']['missed_days'] = 0
                        data['meet_report']['missed_days'] += 1
                        data['meet_report'][date] = 'No'
                    else: 
                        if data['meet_report'].get('active_days','')=='':
                            data['meet_report']['active_days'] = 0
                        data['meet_report']['active_days'] += 1
                        data['meet_report'][date] = 'Yes'
        return analyzed_data            
                        
    def meet_total_time_analysis(analyzed_data):
        print('adding total time spent in meets and calculating averages')
        date = Date()
        grade_list = []
        for user, data in analyzed_data.items():
            day_list = ['monday','tuesday','wednesday','thursday','friday']
            if '@' in user:
                if data['grade'] not in grade_list:
                    grade_list.append(data['grade'])
                if analyzed_data['data'].get('raw_time_in_meet','') =='':
                     analyzed_data['data']['raw_time_in_meet'] = 0
                for day in day_list:
                    if data['meet_report'].get(day + ' time','') == '':
                        continue
                    else:
                        analyzed_data['data']['raw_time_in_meet'] += data['meet_report'][day + ' time']
        analyzed_data['data']['total_time_in_meet'] = date.secs_to_hrs_mins_secs(analyzed_data['data']['raw_time_in_meet'])
        analyzed_data['data']['average_time_in_meet'] = date.secs_to_hrs_mins_secs(analyzed_data['data']['raw_time_in_meet'] / analyzed_data['data']['# of students'])
        analyzed_data['data']['average_time_in_meet_raw'] = analyzed_data['data']['raw_time_in_meet'] / analyzed_data['data']['# of students']
        for grade in grade_list:
            if grade == 'Students':
                continue
            percent = ((analyzed_data['data']['average_time_in_meet_raw_' + grade] / analyzed_data['data']['average_time_in_meet_raw']) - 1)*100
            analyzed_data['data']['%d from school average of average time spent on meets by '+grade] = '{}%'.format(round(percent,2))
        return analyzed_data

    def plan_b_intrs_data_calcs(analyzed_data):
        d = analyzed_data['data']
        j = (((d['total active days in classroom'] + d['total active days by login'] + ['total active days in meet'])/d['# of students'])/5)*100
        d['average % days active'] = '{}%'.format(j)
        analyzed_data['data']['plan_b total'] = 0
        analyzed_data['data']['plan_c total'] = 0
        d['# in moderate alert for plan_b'] = 0
        d['total active days for plan b'] = 0
        d['total active days for plan b no wednesday'] = 0
        d['# in moderate alert for plan_c'] = 0
        d['total time in meets plan b'] = 0
        d['total time in meets plan b no wednesday'] = 0
        for user, data in analyzed_data.items():
            login_log = data['login_log']
            meet_report = data['meet_report']
            class_log = data['classroom_log']
            listy = [login_log, meet_report,class_log]
            if '@' in user:
                if data['grade'] == 'Students':
                    continue
                if data.get('plan_b','') != '':
                    d['plan_b total'] += 1
                    if data['alert'] == 'Moderate':
                        d['# in moderate alert for plan_b'] += 1
                    elif login_log['alert'] == 'Moderate':
                        d['# in moderate alert for plan_b'] += 1
                    elif meet_report['alert'] == 'Moderate':
                        d['# in moderate alert for plan_b'] += 1
                    if data['plan_b'] == 'monday & tuesday':
                        for log in listy:
                            if log['wednesday'] == 'Yes':
                                d['total active days for plan b'] += 1
                            if log['thursday'] == 'Yes':
                                d['total active days for plan b'] += 1
                                d['total active days for plan b no wednesday'] += 1
                            if log['friday'] == 'Yes':
                                d['total active days for plan b'] += 1
                                d['total active days for plan b no wednesday'] += 1
                            d['total time in meets plan b'] += meet_report['wednesday time'] + meet_report['thursday time'] + meet_report['friday time']
                            d['total time in meets plan b no wednesday'] += meet_report['thursday time'] + meet_report['friday time']
                    if data['plan_b'] == 'thursday & friday':
                        for log in listy:
                            if log['wednesday'] == 'Yes':
                                d['total active days for plan b'] += 1
                            if log['monday'] == 'Yes':
                                d['total active days for plan b'] += 1
                                d['total active days for plan b no wednesday'] += 1
                            if log['tuesday'] == 'Yes':
                                d['total active days for plan b'] += 1
                                d['total active days for plan b no wednesday'] += 1
                            d['total time in meets plan b'] += meet_report['wednesday time'] + meet_report['tuesday time'] + meet_report['monday time']
                            d['total time in meets plan b no wednesday'] += meet_report['tuesday time'] + meet_report['monday time']
                else:
                    d['plan_c total'] += 1
                    if data['alert'] == 'Moderate':
                        d['# in moderate alert for plan_c'] += 1
                    elif login_log['alert'] == 'Moderate':
                        d['# in moderate alert for plan_c'] += 1
                    elif meet_report['alert'] == 'Moderate':
                        d['# in moderate alert for plan_c'] += 1
                 
                    

    
    
    date = Date()
    monday = date.nearest_monday()
    print('monday is being anaylzed')
    report_sunday = report_raw(str(date.day_before(monday,1)))
    report_monday = report_raw(str(monday))   
    mon_analyzed_data = report_to_report(report_sunday,report_monday,'mon')

    print('tuesday is being anaylzed')
    report_tuesday = report_raw(str(date.day_after(monday,1)))
    tues_analyzed_data = report_to_data(report_tuesday,mon_analyzed_data,'tues','mon')

    print('wednesday is being anaylzed')
    report_wednesday = report_raw(str(date.day_after(monday,2)))
    wednes_analyzed_data = report_to_data(report_wednesday,tues_analyzed_data,'wednes','tues')

    print('thursday is being anaylzed')
    report_thursday = report_raw(str(date.day_after(monday,3)))
    thurs_analyzed_data = report_to_data(report_thursday,wednes_analyzed_data,'thurs','wednes')

    print('friday is being anaylzed')
    report_friday = report_raw(str(date.day_after(monday,4)))
    fri_analyzed_data = report_to_data(report_friday,thurs_analyzed_data,'fri','thurs')

    add_meet_activity = meet_report(fri_analyzed_data)
    
    final_data = further_data_anaylasis(add_meet_activity,date.day_after(monday,4))
    
    final_data = meet_total_time_analysis(final_data)
    return final_data
    
def report_generator_gspread(analyzed_data):
    date = Date()
    monday = date.nearest_monday()
    gc = gspread.service_account(filename='d:\\code library\\Reports\\Service_account_cred.json')

    sh = gc.create('Report For Week of: {} - {}'.format(monday,date.day_after(monday,4)))
    user_list= ['npadgett','adawson','mfreville','lwalters','rberry','crivera','sjackson']
    for user in user_list:
        user_email = user + '@raleighoakcharter.org'
        sh.share(user_email,perm_type='user', role='writer',email_message='This is the full report for the week of {} to {}'.format(monday,date.day_after(monday,4) ))
    
    #sh.share('npadgett@raleighoakcharter.org', perm_type='user', role='writer',email_message='Hey ya\'ll this is a test of the report generator')
    print('creating spreadsheet')

    sh.add_worksheet(title='Summary Page',rows=67,cols=15)
    sh.add_worksheet(title='Entire School',rows=700,cols=19)

    worksheet_1 = sh.worksheet('Entire School')
    worksheet_0= sh.worksheet('Summary Page')
    worksheet_del = sh.worksheet('Sheet1')
    sh.del_worksheet(worksheet_del)
   
    # formatting for main paige
    n = 1
    list_of_values_main = [{
    'range': 'A1:T{}'.format(n),
    'values': [['Name:','Alert_Level','Teacher:','Grade:','Monday({}): '.format(monday),'Time:','Tuesday({}):'.format(date.day_after(monday,1)),'Time:','Wednesday({}):'.format(date.day_after(monday,2)),'Time:',
    'Thursday({}):'.format(date.day_after(monday,3)),'Time:','Friday({}):'.format(date.day_after(monday,4)),'Time:',
    'Active Days:','Missed Days',
    'Docs Created','Docs Edited','Docs Deleted']]}]
    body = {"requests": []}


    body['requests'].append([{
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": worksheet_1.id,
                "dimension": 'COLUMNS',
                "startIndex": 0,
                "endIndex": 15
            }
        }
    }])
    body['requests'].append([{
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": worksheet_1.id,
                "dimension": 'ROWS',
                "startIndex": 0,
                "endIndex": n}
        }
    }])

    list = []
    for z in range(1,20):
        list.append('')
    for user, data in analyzed_data.items():
        if '@' in user:
            if data['grade'] == 'Students':
                continue
            class_log = data['classroom_log']
            drive_log = data['drive_log']
            login_log = data['login_log']
            meet_log = data['meet_report']
            n += 1 
            list_of_values_main[0]['values'].append([ data['name'], data['alert'],data['teacher'],data['grade'],class_log['monday'], class_log['mon_login'],
            class_log['tuesday'],class_log['tues_login'],class_log['wednesday'],class_log['wednes_login'],class_log['thursday'],
            class_log['thurs_login'],class_log['friday'],class_log['fri_login'],class_log['active_days'],class_log['missed_days'],
            drive_log['total docs created'],drive_log['total docs viewed'],drive_log['total docs edited']])
            list_of_values_main[0]['values'].append(['Plan B: {}'.format(data['plan_b']),login_log['alert'],'Login Data','',login_log['monday'], login_log['mon_login'],
            login_log['tuesday'],login_log['tues_login'],login_log['wednesday'],login_log['wednes_login'],login_log['thursday'],
            login_log['thurs_login'],login_log['friday'],login_log['fri_login'],login_log['active_days'],login_log['missed_days']])
            list_of_values_main[0]['values'].append(['',meet_log['alert'],'Meet Data','Time Spent: {}'.format(meet_log['total_time_in_meets']),meet_log['monday'], date.secs_to_hrs_mins_secs(meet_log['monday time']),
            meet_log['tuesday'],date.secs_to_hrs_mins_secs(meet_log['tuesday time']),meet_log['wednesday'],date.secs_to_hrs_mins_secs(meet_log['wednesday time']),meet_log['thursday'],
            date.secs_to_hrs_mins_secs(meet_log['thursday time']),meet_log['friday'],date.secs_to_hrs_mins_secs(meet_log['friday time']),meet_log['active_days'],meet_log['missed_days']])
            list_of_values_main[0]['values'].append(list)
            body['requests'].append([{
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet_1.id,
                        "startRowIndex": n-1,
                        "endRowIndex": n
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": .8,
                                "green": 1,
                                "blue": .9
                            },
                        }

                    },
                    "fields": "userEnteredFormat(backgroundColor)"
                }
            },
            ])

            n += 1
            body['requests'].append([{
                "mergeCells": {
                    "mergeType": "MERGE_ALL",
                    "range": {
                        "sheetId": worksheet_1.id,
                        "startRowIndex": n-1,
                        "endRowIndex": n,
                        "startColumnIndex": 2,
                        "endColumnIndex": 4
                    }
                }
            },
                {
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet_1.id,
                        "startRowIndex": n-1,
                        "endRowIndex": n
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": .95,
                                "green": 1,
                                "blue": 1
                            }

                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor)"
                }
            }])
            n+=2
            body['requests'].append([
                
                {"repeatCell": {
                    "range": {
                        "sheetId": worksheet_1.id,
                        "startRowIndex": n-1,
                        "endRowIndex": n
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 0,
                                "green": 0,
                                "blue": 0
                            }

                        }

                    },
                    "fields": "userEnteredFormat(backgroundColor)"
                }
                },
                {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet_1.id,
                        "dimension": 'ROWS',
                        "startIndex": n-1,
                        "endIndex":n }
    ,
                "properties": {
                    "pixelSize" : 1
                    },
                "fields": "pixelSize"
                                        }
    }])
    #Add data to Main Summary Page
    x = 1
    list_of_values_summary = [{
        'range': 'A1:O67',
        'values': []}]
    with open('d:\\data\\reports\\TLG.csv','r') as f:
        csvreader = csv.reader(f)
        for row in csvreader:
            listy = []
            for val in row:
                if val in analyzed_data['data']:
                    if analyzed_data['data'][val] == analyzed_data['data']['alert_list']:
                        listy.append(';'.join(analyzed_data['data'][val]))
                    else:
                        listy.append(analyzed_data['data'][val])
                else:
                    listy.append(val)
            list_of_values_summary[0]['values'].append(listy)
    #first value is start row, second is end row, 3rd is start column, and 4th is end column last value is type of border and bolding, 5th is type of border styling seperate from the rest, 6th is a list of
    #postional changes that will affect where a solid thick line is drawn

    df = dict_format_creator()
    body_sum = {"requests": []}
    for value in df.values():
        if value[4] != '':
            style_top = 'SOLID_' + value[4]
            style_bottom = 'SOLID_' + value[4]
            style_right = 'SOLID_' + value[4]
            style_left = 'SOLID_' + value[4]
        else:
             style_top = 'SOLID'
             style_bottom = 'SOLID'
             style_right = 'SOLID'
             style_left = 'SOLID'
        if value[5] != []: 
            c = 1
        if 'left' in value[5]:
            style_left = 'SOLID_THICK'
        if 'right' in value[5]:
            style_right = 'SOLID_THICK'
        if 'top' in value[5]:
            style_top = 'SOLID_THICK'
        if 'bottom' in value[5]:
            style_bottom = 'SOLID_THICK'
        body_sum['requests'].append([{
        "mergeCells": {
                        "mergeType": "MERGE_ALL",
                        "range": {
                            "sheetId": worksheet_0.id,
                            "startRowIndex": value[0],
                            "endRowIndex": value[1],
                            "startColumnIndex": value[2],
                            "endColumnIndex": value[3]
                    }
                }},
        {'updateBorders' : {
                        'range' : {
                            "sheetId": worksheet_0.id,
                            "startRowIndex": value[0],
                            "endRowIndex": value[1],
                            "startColumnIndex": value[2],
                            "endColumnIndex": value[3]
                        },
                        'top' : {
                            'style' : style_top,
                            'color' : {
                                'red': 255,
                                'blue' : 255,
                                'green': 255
                            }
                        },
                        'bottom' : {'style' : style_bottom,
                            'color' : {
                                'red': 255,
                                'blue' : 255,
                                'green': 255}},
                        'left' : {'style' : style_left,
                            'color' : {
                                'red': 255,
                                'blue' : 255,
                                'green': 255}},
                        'right': {'style' : style_right,
                            'color' : {
                                'red': 255,
                                'blue' : 255,
                                'green': 255}


        }}

    }])
    
    body_sum['requests'].append([{
            "mergeCells": {
                            "mergeType": "MERGE_ALL",
                            "range": {
                                "sheetId": worksheet_0.id,
                                "startRowIndex": 4,
                                "endRowIndex": 7,
                                "startColumnIndex": 0,
                                "endColumnIndex": 4
                        }
                    }},
            {'updateBorders' : {
                            'range' : {
                                "sheetId": worksheet_0.id,
                                "startRowIndex": 4,
                                "endRowIndex": 7,
                                "startColumnIndex": 0,
                                "endColumnIndex": 4
                            },
                            'top' : {
                                'style' : 'SOLID_THICK',
                                'color' : {
                                    'red': 255,
                                    'blue' : 255,
                                    'green': 255
                                }
                            },
                            'bottom' : {'style' : 'SOLID_THICK',
                                'color' : {
                                    'red': 255,
                                    'blue' : 255,
                                    'green': 255}},
                            'left' : {'style' : 'SOLID_THICK',
                                'color' : {
                                    'red': 255,
                                    'blue' : 255,
                                    'green': 255}},
                            'right': {'style' : 'SOLID_THICK',
                                'color' : {
                                    'red': 255,
                                    'blue' : 255,
                                    'green': 255}


            }}

    }])

    # begin Formating of data
    x += 1
    n += 1
    body_sum['requests'].append([{
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": worksheet_0.id,
                "dimension": 'COLUMNS',
                "startIndex": 0,
                "endIndex": 15
            }
        }
    }])
    body_sum['requests'].append([{
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": worksheet_0.id,
                "dimension": 'ROWS',
                "startIndex": 0,
                "endIndex": 67}
        }
    }])
    body_sum['requests'].append([{
        "updateDimensionProperties": {
            "range": {
                "sheetId": worksheet_0.id,
                "dimension": 'COLUMNS',
                "startIndex": 0,
                "endIndex":4 }
        ,
        'properties': {
            'pixelSize' : 100
        },
        'fields': 'pixelSize'
    }}])


    for x in range(2,n,4):
        body['requests'].append([{
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet_1.id,
                        "startRowIndex": x-1,
                        "endRowIndex": x,
                        "startColumnIndex":0,
                        "endColumnIndex" : 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "fontSize": 12,
                                "fontFamily" : "Arial"
                            },
                        }

                    },
                    "fields": "userEnteredFormat(textFormat)"
                }
            },
            ])

    gsf.set_frozen(worksheet_1, rows=1)
    rule_1 = gsf.ConditionalFormatRule(
        ranges=[gsf.GridRange.from_a1_range('E1:U', worksheet_1)],
        booleanRule=gsf.BooleanRule(
            condition=gsf.BooleanCondition('TEXT_CONTAINS', ['Yes']),
            format=gsf.CellFormat(textFormat=gsf.textFormat(bold=True), backgroundColor=gsf.Color(0, 1, 0))))
    rule_2 = gsf.ConditionalFormatRule(
        ranges=[gsf.GridRange.from_a1_range('E1:U', worksheet_1)],
        booleanRule=gsf.BooleanRule(
            condition=gsf.BooleanCondition('TEXT_CONTAINS', ['No']),
            format=gsf.CellFormat(textFormat=gsf.textFormat(bold=True), backgroundColor=gsf.Color(1, 0, 0))))

    rule_3 = gsf.ConditionalFormatRule(
        ranges=[gsf.GridRange.from_a1_range('B', worksheet_1)],
        booleanRule=gsf.BooleanRule(
            condition=gsf.BooleanCondition('TEXT_CONTAINS', ['Low']),
            format=gsf.CellFormat(textFormat=gsf.textFormat(bold=True), backgroundColor=gsf.Color(0, 1, 0))))
    rule_4 = gsf.ConditionalFormatRule(
        ranges=[gsf.GridRange.from_a1_range('B', worksheet_1)],
        booleanRule=gsf.BooleanRule(
            condition=gsf.BooleanCondition('TEXT_CONTAINS', ['Moderate']),
            format=gsf.CellFormat(textFormat=gsf.textFormat(bold=True), backgroundColor=gsf.Color(1, 1, 0))))
    rule_5 = gsf.ConditionalFormatRule(
        ranges=[gsf.GridRange.from_a1_range('B', worksheet_1)],
        booleanRule=gsf.BooleanRule(
            condition=gsf.BooleanCondition('TEXT_CONTAINS', ['High']),
            format=gsf.CellFormat(textFormat=gsf.textFormat(bold=True), backgroundColor=gsf.Color(1, 0, 0))))

    rules = gsf.get_conditional_format_rules(worksheet_1)
    rules.append(rule_1)
    rules.append(rule_2)
    rules.append(rule_3)
    rules.append(rule_4)
    rules.append(rule_5)
    rules.save()

    
    worksheet_1.format('B1:U{}'.format(n), {
        'textFormat': {
            'fontSize': 9,
            'fontFamily': 'Arial',
        }
    })
    worksheet_1.format('A1:S1'.format(n), {
        'textFormat': {
            'fontSize': 12,
            'fontFamily': 'Arial',
        }
    })
    worksheet_0.format('A1:U{}'.format(n), {
        'textFormat': {
            'fontSize': 10,
            'fontFamily': 'Arial',
        }
    })

    worksheet_0.format('A8',{
        'verticalAlignment' : 'TOP',
        'textFormat': {
            'fontSize': 12
        }
    })
    
    list_of_values_main[0]['range'] = 'A1:U{}'.format(n)

    worksheet_1.batch_update(list_of_values_main)
    worksheet_0.batch_update(list_of_values_summary)
    sh.batch_update(body)
    sh.batch_update(body_sum)
    worksheet_1.format('A1:U{}'.format(n),{
            'wrapStrategy' : 'WRAP'
        })
    worksheet_0.format('A1:U{}'.format(n),{
            'wrapStrategy' : 'WRAP'
        })
    worksheet_0.format('I26',{
        'borders': {
            'top':{
                'style': 'SOLID_THICK'
            },
            'left': {
                'style':'SOLID_THICK'
            },
            'bottom': {
                'style':'SOLID_THICK'
            }
        }
    })
    worksheet_0.format('E26',{
        'borders': {
            'top':{
                'style': 'SOLID_THICK'
            },
            'left': {
                'style':'SOLID_THICK'
            },
            'bottom': {
                'style':'SOLID_THICK'
            }
        }
    })

    worksheet_0.format('E8:N8',{
            'borders': {
                'top':{
                    'style': 'SOLID_THICK'
                },
                'bottom': {
                    'style':'SOLID_THICK'
                },
                'left': {
                    'style':'SOLID_THICK'
                },
                'right': {
                    'style':'SOLID_THICK'
                }
            }
        })
    worksheet_0.format('E29:O29',{
            'borders': {
                'top':{
                    'style': 'SOLID_THICK'
                },
                'bottom': {
                    'style':'SOLID_THICK'
                },
                'left': {
                    'style':'SOLID_THICK'
                },
                'right': {
                    'style':'SOLID_THICK'
                }
            }
        })
    worksheet_0.format('E45',{
            'borders': {
                'bottom': {
                    'style':'SOLID_THICK'
                },
                'left': {
                    'style':'SOLID_THICK'
                }
        }})
    worksheet_0.format('I45',{
        'borders': {
            'bottom': {
                'style':'SOLID_THICK'
            },
            'left': {
                'style':'SOLID_THICK'
            }
    }})
    worksheet_0.format('O45',{
        'borders': {
            'bottom': {
                'style':'SOLID_THICK'
            },
            'right': {
                'style':'SOLID_THICK'
            }
    }})

def local_report_generator(analyzed_data): 
    listy_prep = []
    listy_dict = {}
    listy_values = []
    date= Date()
    monday = date.nearest_monday()
    friday = str(date.day_after(monday,4))
    for user, data in analyzed_data.items():
        if '@' in user:
            if data['grade'] == 'Students':
                continue
            for key,value in data.items():
                if type(value) == dict:
                    for sec_key,sec_value in data[key].items():
                        if sec_key not in listy_prep:
                            listy_prep.append(sec_key)
                            listy_dict[sec_key] = ''
                if key not in listy_prep:
                    listy_prep.append(key)
                    listy_dict[key] = ''
    with open('D:\\data\\reports\\data_users\\user_data_for_week_of_{} - {}.csv'.format(monday,friday),'w') as f:
        writer = csv.writer(f)
        writer.writerow(listy_prep)
        for user, data in analyzed_data.items():
            listy_dict_user = listy_dict
            if '@' in user:
                if data['grade'] == 'Students':
                    continue
                for key_head,value_head in listy_dict_user.items():
                    for key,value in data.items():
                        if type(value) == dict:
                            for sec_key,sec_value in data[key].items():
                                if sec_key == key_head:
                                    listy_dict_user[key_head] = sec_value
                                    continue
                        if key ==key_head:
                            listy_dict_user[key_head] = value
                for key, value in listy_dict_user.items():
                    listy_values.append(value)
                writer.writerow(listy_values)
                listy_values = []

def main():
    user_list_generator()
    wr= week_report()
    data_keys(wr)
    print(create_class_list(wr))
    #local_report_generator(wr)
    #report_generator_gspread(wr)

if __name__ == '__main__':
    main()
