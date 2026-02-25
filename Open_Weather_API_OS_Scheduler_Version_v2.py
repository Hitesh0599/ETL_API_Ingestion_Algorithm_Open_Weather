#!/usr/bin/env python
# coding: utf-8

# ## 1.0 Importing Libraries

# In[22]:


import requests
import json
import os
from dotenv import load_dotenv
import pandas as pd
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from datetime import datetime
from getpass import getpass
from email import encoders
from IPython.display import HTML, display
import time
import numpy as np


# ## 2.0 Defining Paths, API Key and App Password

# In[23]:


#Defining the url and api_key
url = "https://api.openweathermap.org/data/2.5/weather"
Input_folder = r"C:\Users\MSIKa\Downloads\Python-Projects\Automation Projects\02 Open Weather API\03 Daily Scheduler Version\01 Input Folder"
Output_folder = r"C:\Users\MSIKa\Downloads\Python-Projects\Automation Projects\02 Open Weather API\03 Daily Scheduler Version\02 Excel Output"
Log_folder = r"C:\Users\MSIKa\Downloads\Python-Projects\Automation Projects\02 Open Weather API\03 Daily Scheduler Version\03 Log File"

Keys_folder = r"C:\Users\MSIKa\Downloads\Python-Projects\Automation Projects\02 Open Weather API\03 Daily Scheduler Version\04 Config Folder"
Keys_path = os.path.join(Keys_folder, "keys.env")

load_dotenv(Keys_path)
api_key = os.getenv("Api_Key")
app_password = os.getenv("App_Password")


# ## 3.0 Defining Functions

# ### 3.1 General Functions

# In[24]:


def fn_errors(response):
    error_dict = {
        "400" : "Invalid request format",
        "401" : "Invalid API Key",
        "403" : "Access Permission not given",
        "404" : "City name incorrect",
        "429" : "Too Many Requests",
        "500" : "Internal Server Error",
        "502" : "Bad Gateway",
        "503" : "Service Unavailable"
    }

    code = response.status_code
    return(error_dict.get(str(code), f"Error Code: {code}"))


def fn_creating_excel(folder, column_list, file_name=None):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    for i in range(0, len(column_list)):
        worksheet.cell(row = 1, column = i + 1, value = column_list[i])

    workbook.save(os.path.join(folder, file_name))


# ### 3.2 Weather Data Retrieval Function

# In[25]:


def fn_getting_weather_data_city_name(url, api_key, city_name, unit_type):
    parameters = {
        "q" : city_name,
        "appid" : api_key,
        "units" : unit_type
    }

    
    #Saving the data requested from url into response
    response = requests.get(url, params = parameters)

    #if status is 200 correct fetching so converting the json string to python dict, else failed print
    if response.status_code==200:
        data = response.json()
        return data
    else:
        return(f"Request failed due to {fn_errors(response)}")


# ### 3.3 Saving Data to Excel

# In[26]:


def fn_complex_dictionary_parsing_function(data, key):
    if not key.startswith("["):
        return data[key]
    else:
        key = key.strip("[]")
        keys = key.split("][")

        current = data
        for k in keys:
            k = k.strip("'\"")
            try:
                current = current[int(k)]
            except ValueError:      
                current = current.get(k, "Unknown Location")
        return current
            


def fn_write_weather_metric_to_excel(Output_df, updation_row, output_column_list, data, key_list, i, Unit_Dict, unit_type, min_max_list):
    metric_value = fn_complex_dictionary_parsing_function(data, key_list[i])        

    #for city col where the country coulmn is Unknown Location
    if metric_value == "":
        metric_value = "-"

    
    #when index <=6 the sequence of the key_list and columns in excel are same 
    if i<=6:                
        if key_list[i] == "['main']['temp']":
            Output_df.loc[updation_row, output_column_list[i]] = f"{metric_value}{Unit_Dict[unit_type]["Temp"]}"
        else:
            Output_df.loc[updation_row, output_column_list[i]] = metric_value
            
    #Special case for min and max temp to fetch the temp range
    elif key_list[i] in ["['main']['temp_min']", "['main']['temp_max']"]:
        min_max_list.append(metric_value)
        if len(min_max_list) == 2:
            temp_range = f"{min_max_list[0]} - ({min_max_list[1]}) {Unit_Dict[unit_type]["Temp"]}"
            Output_df.loc[updation_row, output_column_list[i-1]] = temp_range
            
    #after index >6 the sequence of the key_list and columns in excel are not in sync, min and max keys are combined into one single column -> Temp Range
    else:
        if key_list[i] == "['wind']['speed']":
            Output_df.loc[updation_row, output_column_list[i-1]] = f"{metric_value} {Unit_Dict[unit_type] ['Wind Speed']}"
        else:
            Output_df.loc[updation_row, output_column_list[i-1]] = f"{metric_value}%"

    return Output_df



def fn_to_check_excel_presence(folder):
    #Creating initial file list and another empty list for storing non temp files only
    initial_files = os.listdir(folder)
    files = []
    
    #Cleaning the files list by removing the temp files(starts with "~")
    for file in initial_files:
        if not file.startswith("~"):
            files.append(file)

    #Checking if excel present in folder or not
    has_excel_file = 0
    for file in files:
        if file.endswith(".xlsx"):
            has_excel_file = 1
            break
    return has_excel_file




#For idempotent/stateful output updation
def find_existing_weather_record(Output_df, updation_row, city_name, country_code):
    #Making excel updation idempotent i.e. if same country's same city and on same date exists then we update that particular record instead of creating a new one
    #Fetching the country, city and Timestamp
    country = country_code
    city = city_name
    timestamp = datetime.now().strftime("%d/%m/%y")

    mask = (Output_df["Country"] == country) & (Output_df["City"] == city) & (Output_df["Weather Timestamp"] == timestamp)
    if len(Output_df.loc[mask,]) > 0:
        updation_row = Output_df.loc[mask,].index[0]
        
    return updation_row




def fn_saving_final_output(data, Unit_Dict, unit_type, Output_folder, city_name, country_code):
    #Created flag for checking whether output file: present or not
    has_output = fn_to_check_excel_presence(Output_folder)

    #If no previous excel exists, create a new one
    output_column_list = ["Country", "City", "Latitude", "Longitude", "Weather", "Description", "Temperature", "Temp Range", "Humidity", "Wind Speed", "Unit Type", "Weather Timestamp", "Error"]
    if has_output == 0:
        fn_creating_excel(Output_folder, output_column_list, file_name="Weather_Report.xlsx")

    #Loading the workbook df
    Output_df = pd.read_excel(os.path.join(Output_folder, "Weather_Report.xlsx"))
    updation_row = len(Output_df)

    #Making sure that if a country's certain city's data already exist at a certain date, then on next run we are updating that record instead of creating a new one
    updation_row = find_existing_weather_record(Output_df, updation_row, city_name, country_code)

    
    #If error while fetching the data for a particular city
    if not isinstance(data, dict):
        for col in output_column_list:
            Output_df.loc[updation_row, col] = "-"
        Output_df.loc[updation_row, "Country"] = country_code
        Output_df.loc[updation_row, "City"] = city_name
        Output_df.loc[updation_row, "Weather Timestamp"] = datetime.now().strftime("%d/%m/%y")
        #data in this case is a string containing error
        Output_df.loc[updation_row, "Error"] = data
    #No error
    else:
        min_max_list = []
    
        #11 elements to be fetched from api data
        key_list = ["['sys']['country']", "name", "['coord']['lat']", "['coord']['lon']", "['weather'][0]['main']", "['weather'][0]['description']", "['main']['temp']", 
               "['main']['temp_min']", "['main']['temp_max']", "['main']['humidity']", "['wind']['speed']"]
        
        
        for i in range(len(key_list)):
            Output_df = fn_write_weather_metric_to_excel(Output_df, updation_row, output_column_list, data, key_list, i, Unit_Dict, unit_type, min_max_list)
                    
    
        Output_df.loc[updation_row, "Unit Type"] = unit_type 
        Output_df.loc[updation_row, "Weather Timestamp"] = datetime.now().strftime("%d/%m/%y")
        Output_df.loc[updation_row, "Error"] = "-"
    # Output_df.loc[updation_row, "Weather Timestamp"] = datetime.now().strftime("%d/%m/%y %H:%M")

    Output_df.to_excel(os.path.join(Output_folder, "Weather_Report.xlsx"), index = False)
    
            


# ### 3.4 Email-related Functions

# In[27]:


def fn_attaching_email(msg, Output_folder):
    weather_file_path = os.path.join(Output_folder, "Weather_Report.xlsx")
    if os.path.exists(weather_file_path):
        try:
            with open(weather_file_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
    
            encoders.encode_base64(part)
    
            part.add_header(
            'Content-Disposition',  # Tells email client how to handle this part
            f'attachment; filename= {os.path.basename(weather_file_path)}'  # Download as this filename
            )
    
            msg.attach(part)
            return msg
        except Exception as e:
            print(f"Unable to attach the Weather Report due to {e}")
            return None
    else:
        print("The Excel Output file does not exist")
        return None






def fn_email_sending_function(receiver_mail, receiver_name, app_password, Output_folder):
    try:
        #Initializing the sender's mail and mail's headers
        sender_mail = "dagarhitesh99@gmail.com"
    
        msg = MIMEMultipart()
        msg["From"] = sender_mail
        msg["To"] = receiver_mail
        msg["Subject"] = f"Weather Report {datetime.now().strftime("%d-%m-%Y")}"

        Timestamp = datetime.now().strftime("%d/%m %H:%M")
        
        #Creating mail's content
        message = f"""
        <h2> Hello {receiver_name}</h2>
        <h3>Requested Weather Report: {Timestamp}<h3/>
        <p>Please find attached requested weather report</p>
    
        """
    
        msg.attach(MIMEText(message, "html"))

        
        #Attaching the excel to the mail
        msg = fn_attaching_email(msg, Output_folder)

        if not msg is None:
            #Connecting to the Gmail's smtp
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(sender_mail, app_password)
            server.send_message(msg)
            server.quit()
        
    except smtplib.SMTPAuthenticationError:
       print("You may need to check your App Password")
    except Exception as e:
       print(f"Error in sending mail due to {e}")


# ## 4.0 Orchestrator

# In[28]:


#Creating a dictionary of dictionaries(for each unit type)
Unit_Dict = {
    "metric" : {"Temp" : "°C", "Wind Speed" : "m/s"},
    "imperial" : {"Temp" : "°F", "Wind Speed" : "mph"},
    "standard" : {"Temp" : "K", "Wind Speed" : "m/s"}
}

#Searching the weather data for Cities mentioned in the City Names input file
unit_type = "metric"
        
if os.path.exists(os.path.join(Input_folder, "City Names.xlsx")):
    city_names_file = pd.read_excel(os.path.join(Input_folder, "City Names.xlsx"))
    city_names_list = city_names_file["City Name"].tolist()
    country_code_list = city_names_file["Country Code"].tolist()

    for i in range(0, len(city_names_list)):
        country_code = country_code_list[i]
        city_name = city_names_list[i]
        data = fn_getting_weather_data_city_name(url, api_key, city_name, unit_type)
        fn_saving_final_output(data, Unit_Dict, unit_type, Output_folder, city_name, country_code)
    
    display(HTML("<b>Bulk Weather Records have been searched</b>"))

else:
    print("The City Names Input file does not exist")


#Sending the Mails
if os.path.exists(os.path.join(Input_folder, "Receiver Info.xlsx")):
    receiver_mails_file = pd.read_excel(os.path.join(Input_folder, "Receiver Info.xlsx"))
    receiver_names_list = receiver_mails_file["Receiver Name"].tolist()
    receiver_mails_list = receiver_mails_file["Receiver Mail"].tolist()
    

    #Created flag for checking whether log file: present or not
    has_log = fn_to_check_excel_presence(Log_folder)

    log_column_list = ["Receiver Name", "Receiver Mail", "SMTP_Status", "Error", "TimeStamp"]
    if has_log==0:
        fn_creating_excel(Log_folder, log_column_list, file_name="Log_File.xlsx")

    Log_file = pd.read_excel(os.path.join(Log_folder, "Log_File.xlsx"))
    curr_date = datetime.now().strftime("%d/%m/%y")
    
    # Loop through each receiver email one by one
    for i in range(0, len(receiver_mails_list)):
        #Creating conditions to be checked later
        mask1 = (Log_file["Receiver Mail"] == receiver_mails_list[i]) & (Log_file["TimeStamp"] == curr_date)
        mask2 = (Log_file["Receiver Mail"] == receiver_mails_list[i]) & (Log_file["TimeStamp"] == curr_date) & (Log_file["SMTP_Status"] == "Rejected")
        mask3 = (Log_file["Receiver Mail"] == receiver_mails_list[i]) & (Log_file["TimeStamp"] == curr_date) & (Log_file["SMTP_Status"] == "Accepted")

        #If mail not sent at all to the receiver mail, we send the mail
        if len(Log_file.loc[mask1, ]) == 0:
            log_updation_row = len(Log_file)
        #If mail was sent to the receiver mail and it was rejected by SMTP, we redo sending mail
        elif len(Log_file.loc[mask2, ]) > 0:
            log_updation_row = Log_file.loc[mask2, ].index[0]
        #If mail was sent to the receiver mail and it was accepted by SMTP, we skip
        elif len(Log_file.loc[mask3, ]) > 0:
            continue
            
        try:
            fn_email_sending_function(receiver_mails_list[i], receiver_names_list[i], app_password, Output_folder)  
            Log_file.loc[log_updation_row, "SMTP_Status"] = "Accepted"
            Log_file.loc[log_updation_row, "Error"] = ""
            
        except Exception as e:
            print(f"The email couldn't be sent to {receiver_names_list[i]} due to {e}")
            Log_file.loc[log_updation_row, "SMTP_Status"] = "Rejected"
            Log_file.loc[log_updation_row, "Error"] = f"{e}"

        Log_file.loc[log_updation_row, "Receiver Name"] = receiver_names_list[i]
        Log_file.loc[log_updation_row, "Receiver Mail"] = receiver_mails_list[i]
        Log_file.loc[log_updation_row, "TimeStamp"] = curr_date
        
        time.sleep(2)
        
    Log_file.to_excel(os.path.join(Log_folder, "Log_File.xlsx"), index = False)
    display(HTML("<b>The Mails have been sent</b>"))
else:
    print("The Receiver Info Input File does not exist")



# In[ ]:





# # 
