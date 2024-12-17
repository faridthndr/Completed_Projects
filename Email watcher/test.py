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
import csv
line = [] 



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
    last_column_data = ''

    try:
        with open('Query.txt', 'r') as f:
         SQL_Query = f.read()
         
         SQL_Query = SQL_Query.replace('--*',Email )    #-----------------------------------------

        
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
        # print(query_results)

        combined_result = ''
        
        for row in query_results:
            combined_result += '     |   '.join(map(str, row[:-2])) + '\n'

        if len(query_results) > 0:
             ExchangeCode = str(query_results[0][6]).strip()
             PP = str(query_results[0][7]).strip()

        line.append(ExchangeCode)     
        line.append(combined_result)
        line. append(PP)
        # print(last_column_data)  
    except Exception as e:
        error = 3 
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
    


    # Close the database connection
    conn.close()


    #-------------------------------------------------------------------------------------------AZURE-------------- 




def Email_Check():
    creds = None
    global Order_mail 
    Mail_Subject = ''
    global line

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

    results = service.users().messages().list(userId='me', labelIds=['INBOX'], q="is:unread").execute()
    messages = results.get('messages', [])

    if os.path.exists('Sheet_Url.txt'): 
        with open('Sheet_Url.txt', 'r') as u:   #--------------read URL for google sheet
            sheet_url = u.read()
    else:
        pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
        return

    sheet = gc.open_by_url(sheet_url).worksheet('Emails')
   
    if not messages:
        pass
    else:
        row = 1
        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id']).execute()
            # msg = service.users().messages().get(userId='me', id=message['id'], format='raw').execute()

            All_MSG_Body = msg['payload']
            headers = All_MSG_Body['headers']
            Order_mail = 0
            line  = []

            Sent_from = ''
            Sent_to = ''
            Mail_Subject = ''


            for header in headers:
                Titel = header['name']
                Title_Value = header['value']

                print(title)
                print(Title_Value)
                if title.lower() == 'from':
                    name2, Sent_from = parseaddr(Title_Value)
                    # if Sent_from.lower() == 'orders@thndr.app':
                        # Order_mail = 1

                if title.lower() == 'subject':
                   Mail_Subject = Title_Value
                   decoded_subject, encoding = decode_header(Mail_Subject)[0]
                   if encoding:
                      Mail_Subject = decoded_subject.decode(encoding)
                                       
                
                if title.lower() in ['to', 'cc']:
                    name, Sent_to = parseaddr(Title_Value)    
                    if Sent_to.lower() == 'orders@thndr.app' and (Sent_from.lower() != 'orders@thndr.app' or "via thndr orders" in name2.lower()): 
                       Order_mail = 1
                       row += 1
                       line.append(Sent_from)
                       line.append(Mail_Subject)
                       print(Sent_from)

                       if 'parts' in All_MSG_Body :
                         MSG_parts = All_MSG_Body['parts'] 
                         for part in MSG_parts:
                             if part['mimeType'] == 'multipart/alternative':
                                for subpart in part['parts']:
                                    if subpart['mimeType'] == 'text/plain':
                                         data = subpart['body']['data']
                                         decoded_data = base64.urlsafe_b64decode(data).decode('utf-8')
                                         line.append(decoded_data)
                                    elif subpart['mimeType'] == 'text/html':
                                         data = subpart['body']['data']
                                         decoded_txt = base64.urlsafe_b64decode(data).decode('utf-8')
                                         soup = BeautifulSoup(decoded_txt, 'html.parser')
                                         decoded_data2 = soup.get_text()
                                         line.append(decoded_data2)
                             else:
                                      if part['mimeType'] == 'text/plain':
                                         data = part['body']['data']
                                         decoded_data = base64.urlsafe_b64decode(data).decode('utf-8')
                                         line.append(decoded_data)
                                      elif part['mimeType'] == 'text/html':
                                           data = part['body']['data']
                                           decoded_txt = base64.urlsafe_b64decode(data).decode('utf-8')
                                           soup = BeautifulSoup(decoded_txt, 'html.parser')
                                           decoded_data2 = soup.get_text()
                                           line.append(decoded_data2)


                            
            if Order_mail == 1:
                error = Azure_Query(Sent_from)
                sheet.insert_rows([line], row=2) 


                service.users().messages().modify(userId='me', id=message['id'], body={'removeLabelIds': ['UNREAD']}).execute()
                Order_mail = 0

email_keywords = ["Email", "Outlook", "Gmail", "Mail"]
previous_titles = get_window_titles()
print('Start....')



while True:
    time.sleep(1)
    current_titles = get_window_titles()
    new_windows = [title for title in current_titles if title not in previous_titles]
    
    for hwnd, title in new_windows:
        if any(keyword.lower() in title.lower() for keyword in email_keywords):
            print('Email_Check') 
            Email_Check()
            exit()
    previous_titles = current_titles
