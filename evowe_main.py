from pickle import FALSE, TRUE
from altair.vegalite.v4.schema.channels import Tooltip
import pandas as pd
from pathlib import Path
import sqlite3
from sqlite3 import Connection
import streamlit as st
import socket, threading
from PIL import Image
import plotly.express as px
import time
from datetime import datetime
from outlook_plugin import *

URI_EVOLWE_DB = "C:\DOCUMENTS_####_IMPORTANT_NEW\personal\Bits\FourthSemester\ProjectCode\evolwe.db"

def get_connection(path: str):
    """Put the connection in cache to reuse if path does not change between Streamlit reruns.
    NB : https://stackoverflow.com/questions/48218065/programmingerror-sqlite-objects-created-in-a-thread-can-only-be-used-in-that-sa
    """
    print(path)
    return sqlite3.connect(path, check_same_thread=False)

def drop_table(conn: Connection,table_name: str):
    query = "DROP TABLE {0}".format(table_name) 
    conn.execute(query)

def init_db(conn: Connection):
    #drop_table(conn,'evolwe_contact')
    conn.execute(
        """CREATE TABLE IF NOT EXISTS evolwe_contact
            (
                CONTACT_NAME STRING,
                RECURRENCE STRING,
                DURATION STRING,
                LAST_CALLED_TIME STRING
            );"""
    )
    conn.commit()

#def get_data(conn: Connection):
#    df = pd.read_sql("SELECT * FROM evolwe_contact", con=conn)
#    return df

def display_availability_page(conn: Connection):
    begin = dt.datetime.today()
    end = begin + dt.timedelta(days=4)
    cal = get_calendar(begin, end)
    appointments = get_appointments(cal)
    user_data = get_due_data(conn)
    print("user data=")
    print(user_data)
    available_slot = get_available_time_slot(appointments, begin, end)
    call_schedule = get_call_schedule(available_slot, user_data, begin, end)
    st.dataframe(call_schedule)
    
    contact_name = st.text_input("Name:")
    if st.button("Call"):
       now = datetime.now()
       #last_called_time = now.strftime("%d-%m-%y %H:%M:%S")
       last_called_time = now.strftime("%c")
       query = "UPDATE evolwe_contact SET LAST_CALLED_TIME='{0}' WHERE CONTACT_NAME='{1}'".format(last_called_time,contact_name) 
       conn.execute(query)
       conn.commit()
    

def get_data(conn: Connection):
    df = pd.read_sql("SELECT * FROM evolwe_contact", con=conn)
    print("data:")
    print(df)
    return df

def get_due_data(conn: Connection):
    df = pd.read_sql("SELECT * FROM evolwe_contact", con=conn)
    indices_to_drop = []
    for idx, row in df.iterrows():
        recurrence = row['RECURRENCE']
        lst = row['LAST_CALLED_TIME']
        last_called = datetime.strptime(lst, "%c")
        now = datetime.now()
        diff = now - last_called
        diff_seconds = int(diff.total_seconds())
        print("get_due_data : diff seconds=",diff_seconds,"for ", row['CONTACT_NAME'],"recurrence=",recurrence,"idx=",idx)
        no_drop = False
        if recurrence == 'Daily' and diff_seconds > 60: #(24*60*60):
           no_drop = True
        elif recurrence == 'weekly' and diff_seconds > 5*60: # (7*24*60*60):
           no_drop = True
        elif recurrence == 'monthly' and diff_seconds > 10*60: #(30*24*60*60):
           no_drop = True
        
        if not no_drop:
           df.drop(idx,inplace=True)
           #indices_to_drop.append(idx)
           
    #df1 = df.iloc[[indices_to_drop],:]
    print("data:")
    print(df)
    return df

def color_time(val):
    lst = datetime.strptime(val, "%c")
    now = datetime.now()
    diff = now - lst
    diff_seconds = int(diff.total_seconds())
    print("lst={0} now={1} diff={2}".format(lst,now,diff_seconds))
    if diff_seconds < 300:
       color = 'green'
    elif diff_seconds > 300 and diff_seconds < 600:
       color = 'yellow'
    else:
       color = 'red'
    return f'background-color: {color}'

def display_contact_page(conn):
    st.caption("My Contacts:")
    df = get_data(conn)
    df = df.style.applymap(color_time, subset=['LAST_CALLED_TIME'])
    st.dataframe(df)
    contact_name = st.text_input("Name:")
    recurrence = st.text_input("Recurrence:")
    duration = st.text_input("Duration:")
    #last_called_time = st.text_input("Last Called Time:")
    st.caption("Note:Text entered are case sensitive")

    if st.button("Add/Update"):
        print("contact_name={0}".format(contact_name))
        query = "SELECT DURATION FROM evolwe_contact WHERE CONTACT_NAME='{0}'".format(contact_name)
        rows = conn.execute(query)
        record_present = False
        for row in rows:
            now = datetime.now()
            #last_called_time = now.strftime("%d-%m-%y %H:%M:%S")
            last_called_time = now.strftime("%c")
            query = "UPDATE evolwe_contact SET CONTACT_NAME='{0}', RECURRENCE='{1}', DURATION='{2}', LAST_CALLED_TIME='{3}'".format(contact_name,recurrence,duration,last_called_time)
            print("update query= ",query)
            record_present = True
        if not record_present:
            now = datetime.now()
            #last_called_time = now.strftime("%m/%d/%Y %H:%M:%S")
            last_called_time = now.strftime("%c")
            query = "INSERT INTO evolwe_contact (CONTACT_NAME, RECURRENCE, DURATION, LAST_CALLED_TIME) VALUES ('{0}','{1}', '{2}','{3}')".format(contact_name,recurrence,duration,last_called_time)
        conn.execute(query)
        conn.commit()
    print("contact name = {0} Recurrence = {1} Duration = {2}".format(contact_name,recurrence,duration))
    

def display_data(conn: Connection):
    df = get_data(conn)
    if st.button("Refresh"):
       df = get_data(conn)
    
    if st.button("Clear"):
       pass

    #df.style.applymap(color_ss, subset=['WL_SECURITY_SCORE'])
    
    #st.dataframe(df.style.applymap(color_ss, subset=['SECURITY_SCORE','WL_SECURITY_SCORE']))
    
    #fig = px.bar(df, x="MAC_ADDRESS", y=["SECURITY_SCORE", "WL_SECURITY_SCORE"], barmode='group', height=500)
    #st.bar_chart(df) # if need to display dataframe
    #st.plotly_chart(fig)

    #st.vega_lite_chart(df,{'mark':{'type':'circle','tooltip':True},
    #                       'encoding':{
    #                        'x': {'field':'SECURITY_SCORE','type':'quantitative'},
    #                        'y': {'field': 'WL_SECURITY_SCORE','type':'quantitative'},
    #                        'tooltip': {'field': 'MAC_ADDRESS'},
    #                        'description': {'field': 'SECURITY_SCORE'}
    #                       }})
    #st.vega_lite_chart()
    

def build_sidebar(conn: Connection):
    st.sidebar.title("VR4EVR")
    st.sidebar.subheader("EvolWE")
    sdbar_select = st.sidebar.radio("Details",
                                    ["My Contacts", "My Free Slots"])
    image = Image.open('together2.jpg')
    st.image(image, caption='Deployment',use_column_width=True)
    if sdbar_select == "My Contacts":
        display_contact_page(conn)
    if sdbar_select == "My Free Slots":
        display_availability_page(conn)
    
def main():
    conn = get_connection(URI_EVOLWE_DB)
    init_db(conn)
    build_sidebar(conn)

if __name__ == "__main__":
    main()