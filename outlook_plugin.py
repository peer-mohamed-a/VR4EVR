import datetime as dt
import pandas as pd
import win32com.client
import numpy as np
import pythoncom 

def get_calendar(begin,end):
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    return calendar

def get_appointments(calendar):
    appointments = [app for app in calendar]    
    #cal_subject = [app.subject for app in appointments]
    #cal_start_date = [app.start.strftime('%m-%d-%Y') for app in appointments]
    #cal_start_time = [app.start.strftime('%H:%M') for app in appointments]
    #cal_end_date = [app.end.strftime('%m-%d-%Y') for app in appointments]
    #cal_end_time = [app.end.strftime('%H:%M') for app in appointments]
    #cal_body = [app.body for app in appointments]
    #cal_day = [app.start.strftime('%a') for app in appointments]
    #df = pd.DataFrame({'subject': cal_subject, 'start': cal_start, 'end': cal_end, 'body': cal_body})
    #df = pd.DataFrame({'start_date': cal_start_date, 'start_time': cal_start_time, 'end_date': cal_end_date, 'end_time': cal_end_time, 'day':cal_day})
    cal_start = [app.start.strftime('%m-%d-%Y %H:%M') for app in appointments]
    cal_end = [app.end.strftime('%m-%d-%Y %H:%M') for app in appointments]
    cal_day = [app.start.strftime('%a') for app in appointments]
    df = pd.DataFrame({'start': cal_start, 'end': cal_end, 'day': cal_day})
    return df

def get_available_time_slot(appointments, begin, end):
    available_time_start = []
    available_time_end = []
    tmp_appointments = appointments.copy(True)
    remaining_appointments = len(tmp_appointments)
    
    day_increment = dt.timedelta(days=1)
    today = begin
    while today <= end:
        #extract day and date for today
        date = today.strftime("%m-%d-%Y")
        day = today.strftime("%a")	
        
        if day == 'Sat' or day == 'Sun':
            #free time on weekend starts at 8 am 
            current_free_time = date + " 8:00"
        else:
            #free time on weekday starts at 4 pm (i.e) 16
            current_free_time = date + " 16:00"
        free_time_end = date + " 22:00"
        dt_today = dt.datetime.strptime(date, "%m-%d-%Y")
        dt_end_free_time = dt.datetime.strptime(free_time_end, "%m-%d-%Y %H:%M")
        
        #iterate day by day and find the available slot
        if remaining_appointments > 0:
            for index, row in tmp_appointments.iterrows():
                #remove the iterated entries from the tmp_appointments	
                tmp_start_time = row['start'][0:10]
                dt_tmp_day = dt.datetime.strptime(tmp_start_time, "%m-%d-%Y")
                dt_current_free_time = dt.datetime.strptime(current_free_time, "%m-%d-%Y %H:%M")
                diff = dt_current_free_time - dt_end_free_time
                if diff.total_seconds() > 0:
                    continue
                diff = dt_tmp_day - dt_today
                if diff.total_seconds() > 0:
                    #dt_tmp_day is on future day, add the remaining slot of today
                    available_time_start.append(dt_current_free_time)
                    available_time_end.append(dt_end_free_time)
                    break
                elif diff.total_seconds() == 0:
                    #both are on same day, iterate and add the free slot of today
                    dt_meeting_start_time = dt.datetime.strptime(row['start'], "%m-%d-%Y %H:%M")
                    dt_meeting_end_time = dt.datetime.strptime(row['end'], "%m-%d-%Y %H:%M")
                    #case 1 : meeting started before 4pm and ended before 4pm
                    #case 2 : meeting started before 4pm and continued till 4+ pm        
                    #case 3 : meeting started after 4pm
            		
                    diff = dt_current_free_time - dt_meeting_start_time
                    diff2 = dt_current_free_time - dt_meeting_end_time
                    if diff.total_seconds() > 0:
                        #meeting started before free time
                        diff = dt_current_free_time - dt_meeting_end_time 
                        if diff.total_seconds() > 0:
                            #meeting started before free time and ended before free time, skip it
                            tmp_appointments.drop(index)
                            remaining_appointments -= 1
                            continue
                        else:
                            #case 3: meeting started before free time and ended after free time
                            current_free_time = row['end']
                            tmp_appointments.drop(index)
                            remaining_appointments -= 1
                            continue
                    else:
                        #meeting started after free time
                        available_time_start.append(dt_current_free_time)
                        diff = dt_end_free_time - dt_meeting_start_time
                        if diff.total_seconds() < 0:
                            available_time_end.append(dt_end_free_time)
                        else:
                            available_time_end.append(dt_meeting_start_time)
                        current_free_time = row['end']
                        tmp_appointments.drop(index)
                        remaining_appointments -= 1
        if remaining_appointments <= 0:
            dt_current_free_time = dt.datetime.strptime(current_free_time, "%m-%d-%Y %H:%M")
            available_time_start.append(dt_current_free_time)
            available_time_end.append(dt_end_free_time)						
        today += day_increment
    #print (available_time_start)
    #print (available_time_end)
    df = pd.DataFrame({'start':available_time_start, 'end':available_time_end})
    return df

def get_user_db():
    data = {
        "name": ["abc", "def", "ghi"],
        "duration": [50, 40, 45],
        "priority": [2,3,1]
    }
    df = pd.DataFrame(data)
    return df

def get_call_schedule(available_slot, user_db, begin, end):
    call_schedule_start = []
    call_schedule_end = []
    call_schedule_name = []
    call_schedule_duration = []
    
    tmp_available_slot = available_slot.copy(True)
    max_slot = len(tmp_available_slot)
    idx = 0
    day_increment = dt.timedelta(days=1)
    today = begin
    max_db = len(user_db)
    print("user db=")
    print(user_db)
    while today <= end:
	    #increment day by day 
        #print ("iterating {0}  {1}  {2}".format(today, begin, end))
        usr_db = user_db.copy(True)
        #usr_db = usr_db.sort_values(by=['priority'])
        db_idx = 0
        usr_db_remaining = max_db
        date = today.strftime("%m-%d-%Y")
        dt_today = dt.datetime.strptime(date, "%m-%d-%Y")
        available_slot_today_start = []
        available_slot_today_end = []
        #get available_slot for a day
        
        for idx, row in tmp_available_slot.iterrows():
            print(idx)
            print(tmp_available_slot['start'][idx])
            tmp_start_time = str(tmp_available_slot['start'][idx])[0:10]
            dt_tmp_day = dt.datetime.strptime(tmp_start_time, "%Y-%m-%d")
            diff = dt_tmp_day - dt_today
            if diff.total_seconds() == 0:
                available_slot_today_start.append(tmp_available_slot['start'][idx])
                available_slot_today_end.append(tmp_available_slot['end'][idx])
            elif diff.total_seconds() > 0:
                break
            else:
                tmp_available_slot.drop(idx)
            #idx += 1
            #print ("idx: {0}/{1}".format(idx, max_slot))			
            #if idx >= max_slot:
            #    break
        for i in range(0, len(available_slot_today_start)):
            #print ("{0}  start:{1} end:{2} db_idx:{3}/{4}".format(date, available_slot_today_start[i], available_slot_today_end[i], db_idx, max_db))
            #check any calls can be fit into the slot
            available_start_time = dt.datetime.strptime(str(available_slot_today_start[i]), "%Y-%m-%d %H:%M:%S")
            available_end_time = dt.datetime.strptime(str(available_slot_today_end[i]), "%Y-%m-%d %H:%M:%S")
            #available_start_time = dt.datetime.strftime((available_slot_today_start[i]), "%c")
            #available_end_time = dt.datetime.strftime((available_slot_today_end[i]), "%c")
            for db_idx, row in usr_db.iterrows():			
                #print("idx {0} {1} db {2} {3}".format(i, len(available_slot_today_start), db_idx, max_db))
                #print(usr_db)
                duration = usr_db['DURATION'][db_idx]
                duration = int(duration)
                #print("available {0} {1} duration {2}".format(available_start_time, available_end_time, duration))
                diff = (((available_end_time - available_start_time).total_seconds())/60) - duration
                if diff >= 0:
                    usr_db_remaining = usr_db_remaining - 1
                    call_schedule_start.append(dt.datetime.strftime(available_start_time,"%c"))
                    available_start_time = available_start_time + dt.timedelta(minutes=int(duration))
                    call_schedule_end.append(dt.datetime.strftime(available_start_time,"%c"))
                    call_schedule_name.append(usr_db['CONTACT_NAME'][db_idx])
                    call_schedule_duration.append(int(duration))
                    usr_db = usr_db.drop(db_idx)
        today += day_increment	
    df = pd.DataFrame({'start':call_schedule_start, 'end':call_schedule_end, 'CONTACT_NAME':call_schedule_name, 'DURATION':call_schedule_duration})
    return df

#begin = dt.datetime.today()
#end = begin + dt.timedelta(days=4)
#cal = get_calendar(begin, end)
#appointments = get_appointments(cal)
#available_slot = get_available_time_slot(appointments, begin, end)
#print (available_slot)
#user_db = get_user_db()
#call_schedule = get_call_schedule(available_slot, user_db, begin, end)
#print(call_schedule)