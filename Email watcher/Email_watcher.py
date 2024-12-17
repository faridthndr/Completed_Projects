import os
import pickle
import base64
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from email.utils import parseaddr
import time
import win32gui
import gspread
from bs4 import BeautifulSoup
import email
from email.header import decode_header
import pandas as pd
import pyodbc
import chardet
import pygetwindow as gw

import csv
line = [] 
line2 = []
service = 0
gc = 0

Order_mail = 0


SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.modify']

def get_window_titles():
    def enum_window_titles(hwnd, titles):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            titles.append((hwnd, win32gui.GetWindowText(hwnd)))
        return True

    titles = []
    win32gui.EnumWindows(enum_window_titles, titles)
    return titles


def Azure_Query(Email):
    
   #-------------------------------------------------------------------------------------------AZURE-------------- 
    
    server = "10.70.4.100"
    database = "master"
    username = "farid.shawky@thndr.app"
    password = "8mrM7dZF@5LP"
    global line
    global line2
    last_column_data = ''
    ExchangeCode = ''
    UnifiedCode = ''

    try:
        with open('Query.txt', 'r') as f:
         SQL_Query = f.read()
         
         SQL_Query = SQL_Query.replace('--*',Email )    #-----------------------------------------

        with open('Query2.txt', 'r') as m:
         SQL_Query2 = m.read()
         
         SQL_Query2 = SQL_Query2.replace('--*',Email )
        
    except Exception as e:
        error = 1
        return error

    # Establish connection to Azure SQL Database
    try:
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        # conn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        conn = pyodbc.connect(conn_str)
    except Exception as e:
        error = 2
        return error

    # Execute the SQL query and fetch the results
    try:
        conn.commit()
        cursor = conn.cursor()
        cursor.execute(SQL_Query)
        query_results = cursor.fetchall()

        combined_result = ''
        # header = [desc[0] for desc in cursor.description]
        # combined_result += ' | '.join([col.center(10) for col in header[:-2]]) + '\n'

        
        for row in query_results:
            combined_result += ' | '.join([str(val).center(10) for val in row[:-3]]) + '\n'



        if len(query_results) > 0:
             ExchangeCode = str(query_results[0][6]).strip()
             UnifiedCode = str(query_results[0][7])
             PP = str(query_results[0][8]).strip()

        line.append(ExchangeCode)
        line.append(UnifiedCode)
        line.append(combined_result)
        line. append(PP)
    except Exception as e:
        error = 3 
        print('error in Azure_Query')
        return error

    
    try:
        # conn.commit()
        # cursor = conn.cursor()
        cursor.execute(SQL_Query2)
        query_results2 = cursor.fetchall()
        combined_result2 =''
        # header = [desc[0] for desc in cursor.description]
        # combined_result2 += ' | '.join([col.center(10) for col in header]) + '\n'


        for row in query_results2:
            combined_result2 += ' | '.join([str(val).center(10) for val in row]) + '\n'

        line.append(combined_result2)
        
    except Exception as e:
        error = 3 
        print(f'error {e}')
        return error
        
    filename = 'Q_results.csv'
    with open(filename, 'w', newline='', encoding='utf-16') as csvfile:
        writer = csv.writer(csvfile)
        # Get the column names from the cursor's description
        # column_names = [column[0] for column in cursor.description]
        # column_names.append("Eligibility")
        # column_names.append("Action")
        

        # Write each row of the query results
        writer.writerows(query_results)


    csvfile.close()   
    
    print('assss')

    # Close the database connection
    conn.close()


    #-------------------------------------------------------------------------------------------AZURE-------------- 


creds = None
Mail_Subject = ''
decoded_data2 = ''

if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

if not creds or not creds.valid:    
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('gmail', 'v1', credentials=creds)
gc = gspread.authorize(creds)

user_id = 'me'  
topic_name = 'projects/email-watcher-422706/topics/emailnotify'

request_body = {  
    'topicName': topic_name  
}  


watch_response = service.users().watch(userId='me', body={'labelIds': ['INBOX'],'topicName': topic_name}).execute()  
last_history_id =  watch_response['historyId'] 
print(last_history_id) 




def Email_Check():
    global service
    global line
    global line2
    """
     creds = None
    global Order_mail 
    Mail_Subject = ''
    decoded_data2 = ''
    global line
    global service
    global  gc

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:    
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    user_id = 'me'  
    topic_name = 'projects/email-watcher-422706/topics/emailnotify'

    request_body = {  
        'topicName': topic_name  
    }  
    
    try:  
        response = service.users().watch(userId=user_id, body=request_body).execute()  
        print('Watch response: ', response)  
    except Exception as e:  
        print(f'Error setting up watch: {e}')  




    gc = gspread.authorize(creds)
    """
    results = service.users().messages().list(userId='me', labelIds=['INBOX'], q="is:unread").execute()
    messages = results.get('messages', [])

    if os.path.exists('Sheet_Url.txt'): 
        with open('Sheet_Url.txt', 'r') as u:   #--------------read URL for google sheet
            sheet_url = u.read()
    else:
        pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
        return

   
    sheet = gc.open_by_url(sheet_url).worksheet('Emails')
    

    # sheet.update_cell(1, 1, "Updated Value")
   
    if not messages:
        pass
    else:
        row = 1
        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id']).execute()
            All_MSG_Body = msg['payload']
            headers = All_MSG_Body['headers']
            Order_mail = 0
            line  = []
            Sent_from = ''
            Sent_to = ''
            Mail_Subject = ''
            decoded_data =''
            decoded_data2 = ''

            for header in headers:
                # print(headers)
                Titel = header['name']
                Title_Value = header['value']

                if Titel.lower() == 'from' :
                    name2 , Sent_from = parseaddr(Title_Value)
                    if Sent_from.lower() =='orders@thndr.app' and "via thndr orders" not in name2.lower():
                       service.users().messages().modify(userId='me', id=message['id'], body={'removeLabelIds': ['UNREAD']}).execute()


                if Titel.lower() == 'subject':
                      Mail_Subject = Title_Value

                      decoded_subject, encoding = decode_header(Mail_Subject)[0]
                      if encoding:
                         Mail_Subject = decoded_subject.decode(encoding)
                                     


                if Titel.lower() in ['to', 'cc']:
                    name, Sent_to = parseaddr(Title_Value)    
                    if Sent_to.lower() == 'orders@thndr.app' and (Sent_from.lower() != 'orders@thndr.app' or "via thndr orders" in name2.lower()): 
                        Order_mail = 1
                        row += 1
                        line.append(message['id'])
                        line.append(Sent_from)
                        line.append(Mail_Subject)
                      

                        if 'parts' in All_MSG_Body:
                            MSG_parts = All_MSG_Body['parts']
                            for part in MSG_parts:
                                if part['mimeType'] == 'multipart/alternative':
                                    for subpart in part['parts']:
                                        if subpart['mimeType'] == 'text/plain':
                                            data = subpart['body']['data']
                                            decoded_data = base64.urlsafe_b64decode(data).decode('utf-8')
                                            line.append(decoded_data)
                                            # print(decoded_data)
                                        elif subpart['mimeType'] == 'text/html':
                                            data = subpart['body']['data']
                                            decoded_txt = base64.urlsafe_b64decode(data).decode('utf-8')
                                            soup = BeautifulSoup(decoded_txt, 'html.parser')
                                            decoded_data2 = soup.get_text()
                                            # line.append(decoded_data2)
                                            # print(decoded_data2)
                                else:
                                    if part['mimeType'] == 'text/plain':
                                        data = part['body']['data']
                                        decoded_data = base64.urlsafe_b64decode(data).decode('utf-8')
                                        # line.append(decoded_data)
                                        # print(decoded_data)
                                    elif part['mimeType'] == 'text/html':
                                        data = part['body']['data']
                                        decoded_txt = base64.urlsafe_b64decode(data).decode('utf-8')
                                        soup = BeautifulSoup(decoded_txt, 'html.parser')
                                        decoded_data2 = soup.get_text()
                                        line.append(decoded_data2)
                                        # print(decoded_data2)
                            # line.append(decoded_data2)
                                               
            if  Order_mail == 1 : 
                # print(Order_mail)
                error = Azure_Query(Sent_from)
                print(Sent_from)
                sheet.insert_rows([line], row=2) 
                

                service.users().messages().modify(userId='me', id=message['id'], body={'removeLabelIds': ['UNREAD']}).execute()
                Order_mail = 0
                
                
                        
# email_keywords = ["Email", "Outlook", "Gmail", "Mail"]
# previous_titles = get_window_titles()
# email_keywords = ["Email", "Outlook", "Gmail", "Mail", "Notification", "thndr","inbox","Orders"]


# previous_titles = get_window_titles()

print('Start....')

while True:
    # exit()
    time.sleep(5)
    try:  
        history_response = service.users().history().list(userId='me', startHistoryId=last_history_id).execute()  
        changes = history_response.get('history', [])  
        
        if changes: 
            Email_Check() 
            print("New messages detected:")  

        else:  
            # print("No new messages.")
            pass  
    except Exception as  error:  
        print(f'An error occurred: {error}') 
    last_history_id =  history_response['historyId'] 
    # print(last_history_id)
    exit()
    