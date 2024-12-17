from oauth2client.service_account import ServiceAccountCredentials
import gspread
import os
from tkinter import messagebox
import csv
import pyodbc



def send_slack_notification(query_results):  
    Slack_url = 'https://hooks.slack.com/services/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' 
    # Slack_url =  'https://hooks.slack.com/services/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'   
    csv_data = []   
    csv_data.append("  POR has been excuted  ")  
    
    
    csv_content = "\n\n".join(csv_data)    
    
    payload = {  
        "text": f"Warning: Number of rejection messages in the last inquiry is:\n```\n{csv_content}\n```"  
    }  
    
    headers = {  
        'Content-Type': 'application/json'  
    }  
    
    response = requests.post(Slack_url, json=payload, headers=headers)  

    if response.status_code != 200:  
        error = 11   
        return error  




def syc():
        
    csv_file_path = 'All_Users_data.csv'



    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_file = 'creds.json'
    google_jsonfile = 'googlekey.json'

    creds = None
    if os.path.exists(creds_file):         
        with open('creds.json', 'r') as f:  
            creds_json = f.read()
            creds = ServiceAccountCredentials.from_json(creds_json)     
    else:
        if not os.path.exists(google_jsonfile):     
            error = 3
        else:
            creds = ServiceAccountCredentials.from_json_keyfile_name(google_jsonfile, scope)   
            with open(creds_file , 'w') as f:
                f.write(creds.to_json())

    client = gspread.authorize(creds)


    if os.path.exists('Sheet_Url.txt'): 
            with open('Sheet_Url.txt', 'r') as u:   
                sheet_url = u.read()

            #Access the Google Sheet
            sheet = client.open_by_url(sheet_url).worksheet('MCDR Alert')
            
    else:
        messagebox.showinfo('Error','Can not find Sheet URL')

    # Find the last row with values
    last_row = len(sheet.col_values(1)) + 1
        
    header =sheet.row_values(1)
    # ID_index = header.index('New UID')
    # ID_Number_index = header.index('New ID Number')
    # google_rows = sheet.get_all_values()
    google_rows = sheet.get_all_values()




    ##-------------------------------------------------------------------------------------AZURE----------------------------------- 
    # Azure SQL Database connection details

    try:
        with open ('Alert_Query.txt','r') as m:
            Alert_Query  = m.read()
            # Alert_Query = Alert_Query.replace('--*',f'-{Timer}' )
            

        
    except Exception as e:
        messagebox.showinfo("Error","Error")
        error = 1
        return e

    # Establish connection to Azure SQL Database
    try:
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        conn = pyodbc.connect(conn_str)
    except Exception as e:
        messagebox.showinfo("Error","Error")
        return e

    # Execute the SQL query and fetch the results
    conn.commit()
    cursor = conn.cursor()
    cursor.execute(Alert_Query)
    query_results = cursor.fetchall()
    print(query_results)


