
import requests
from tkinter import messagebox
import pyodbc
import time 



user_name =''
password = '' 
server = ''
database = ''
symbol_code = '' 
timer = ''


def send_slack_notification( ):
    global symbol_code  
    Slack_url = 'https://hooks.slack.com/services/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' 
   
    payload = {  
        "text": f"Orders for the symbol code ({symbol_code}) have been executed:\n\n"  
    }  
    
    headers = {  
        'Content-Type': 'application/json'  
    }  
    
    response = requests.post(Slack_url, json=payload, headers=headers)  

    if response.status_code != 200:
        messagebox.showinfo("Error","Error in Alert Function")
        return False  




def Query_Fetch():

    global user_name, password, server ,database ,symbol_code 
    
    try:
        with open('Query.txt', 'r') as f:
            SQL_Query = f.read()
            SQL_Query = SQL_Query.replace('--*',symbol_code )    #----------------------------------------- 


        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={user_name};PWD={password}'
        conn = pyodbc.connect(conn_str)
        conn.commit()
        cursor = conn.cursor()
        cursor.execute(SQL_Query)
        query_results = cursor.fetchall()

    except Exception as d: 
        messagebox.showinfo("Error",d)     
        return False

    conn.close()
    
    return query_results



while True:


    try:

        config = {} 
        config_file_path = 'config.txt' 
    
        with open(config_file_path, 'r') as file:  
            for line in file:  
                line = line.strip()  
                if '=' in line:  
                    key, value = line.split('=', 1)  
                    key = key.strip()  
                    value = value.strip()  
                    config[key] = value  
        print(config)
    except Exception as e:
        messagebox.showinfo("Error","Error in config file ")  
      

    user_name = config.get('user name')  
    password = config.get('password')  
    server = config.get('server')  
    database = config.get('Database')  
    symbol_code = config.get('symbol code')  
    timer = config.get('timer')  

    if not Query_Fetch():
        pass
        
    else:
        if not send_slack_notification() :
            pass
    for i  in range(int(timer)):
        time.sleep(1)
        print(f'Remaining Time {int(timer)-i}')
    

