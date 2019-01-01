# -*- coding: utf-8 -*-
"""
Created on Fri Jan 19 16:50:46 2018

@author: orionis
"""

import pandas as pd #To make dataframes from HTML
import numpy as np #To deal with NaNs
import re #To parse and extract necessary information from the strings
import httplib2 #To establish connection with Google Calendar
import os #To get the credential files

#Following are necessary imports for working of Google calendar API
from apiclient import discovery 
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import datetime

df = pd.read_html("University of Petroleum & Energy Studies (23042018).html")

working_df = df[3].iloc[:-1, 4:13]

working_df.replace(np.nan, 'missing', inplace = True) #Replace all NaN values by 'missing'


#the whole dataframe in a list
sub_vals = []

for row in working_df.values:
    sub_vals.append(row)
sub_vals = [[stuff for stuff in sub_vals[i] if #Using list comprehension to get only those values which are not NULL 
             stuff != 'missing'] for i in range(len(sub_vals))]

#Lists to store the sub names, dates, start and endtimes
names = []
dates = []
starttimes = []
endtimes = []

for counter in range(len(sub_vals)):
    for things in sub_vals[counter]:
        #Temp lists
        subnames = []
        room = []
        date = []
        time = []
        #Since names in schedule appear before the first parenthesis
        for subname in things:
            if subname == '(':
                break
            subnames.append(subname)
        name = "".join(subnames).rstrip()
        
        #Room number appears after the time
        rno_txt = re.split(r'\d+:\d{2}', things)
        for number in rno_txt[2]:
            if number == '(':
                break
            room.append(number)
        rno = "".join(room).rstrip()
        
        #Appending to list the room number and the subject name
        names.append(rno + " " + name)
        
        #Time appears after the word Schedule
        time_txt = things.partition('Schedule: ')[2]    
        for ti in time_txt:
            if ti == ' ':
                break
            time.append(ti)
        tim = "".join(time)
        
        #Appending the start times to the list
        starttimes.append(tim)
        
        end_time = re.findall(r'\d+:\d{2}', things)
        
        #Appending end times to the list
        endtimes.append(end_time[1])
                    
           
dateframe = df[3][df[3].columns[0]]
dateframe.drop(dateframe.index[5], inplace = True)

#Getting the dates
for counter in range(len(dateframe)):
    dat = []
    for date in dateframe[counter]:
        if date == "(":
            break
        dat.append(date)
    date = "".join(dat).rstrip()
    for number in range(len(sub_vals[counter])):
        dates.append(date)
    
#Adding seconds to the times, requirement of Google Calendar API
starttimes = [starttimes[index] + ':00' for index in range(len(starttimes))]
endtimes = [endtimes[index] + ':00' for index in range(len(endtimes))]
#Changing format of dates, requirement of Google Calendar API
dates = [datetime.datetime.strptime(dates[index], '%d-%m-%Y').strftime('%Y-%m-%d') 
        for index in range(len(dates))]

#Lists to store final times in the required format
final_start_time = []
final_end_time = []

for _ in range(len(dates)):
    for dt, st, et in zip(dates, starttimes, endtimes):        
        final_start_time.append(dt + 'T' + st + '+05:30')
        final_end_time.append(dt + 'T' + et + '+05:30')            

#A dictionary of dictionaries containing the events.
event_dict = {}

for suffix in range(len(names)):
    event_dict['event{0}'.format(suffix)] = {'start': {'timeZone': 'GMT +05:30'},
                                                 'end': {'timeZone': 'GMT +05:30'},
                                                 'reminders': {'useDefault': False}
                                                 }    
    

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'CalTest'

def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

#Uploading the events to Google Calendar
for counter in range(len(names)):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)
    event_dict['event{0}'.format(counter)]['summary'] = names[counter]
    event_dict['event{0}'.format(counter)]['start']['dateTime'] = final_start_time[counter]
    event_dict['event{0}'.format(counter)]['end']['dateTime'] = final_end_time[counter]
    event_dict['event{0}'.format(counter)] = service.events().insert(calendarId='primary',
               body=event_dict['event{0}'.format(counter)]).execute()
    print('Event created: %s' % (event_dict['event{0}'.format(counter)].get('htmlLink')))