import os
import pickle
import base64
import re
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from email.utils import parseaddr
import time
import win32gui
import gspread
from bs4 import BeautifulSoup




Order_mail = 0



# SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
# SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/gmail.modify']
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.modify']




def get_window_titles():
    def enum_window_titles(hwnd, titles):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):

            titles.append((hwnd, win32gui.GetWindowText(hwnd)))
        return True

    titles = []
    win32gui.EnumWindows(enum_window_titles, titles)
    return titles



def Email_Check():
    creds = None
    global Order_mail 
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

            Sent_from = ''
            Sent_to = ''
            Mail_Subject = ''

            for header in headers:
                # print(headers)
                Titel = header['name']
                Title_Value = header['value']

                if Titel.lower() == 'subject':
                      Mail_Subject = Title_Value


                if Titel.lower() == 'from' :

                    name2 , Sent_from = parseaddr(Title_Value)
                    # Sent_from = Title_Value
                    if  Sent_from.lower() == 'orders@thndr.app':
                        # print(Mail_Subject)
                        # print(Sent_from)
                        Order_mail = 1 
                        if 'parts' in All_MSG_Body:  
                              MSG_parts = All_MSG_Body['parts']
                              for part in MSG_parts:
                                    if part['mimeType'] == 'text/plain':
                                        data = part['body']['data']
                                        # decoded_data = base64.urlsafe_b64decode(data).decode('utf-8')
                                        Order_mail = 1
                                        # print(decoded_data) 

                
                

                if Titel.lower() == 'to' or Titel.lower() == 'cc':
                   # Sent_to =  Title_Value
                   name, Sent_to = parseaddr(Title_Value)
                   # print(name)
                   # print(Sent_to)

                   
                   if Sent_to.lower() == 'orders@thndr.app' and ( Sent_from.lower() != 'orders@thndr.app' or "via thndr orders" in name2.lower()) : 
                          # print(Mail_Subject)
                          Order_mail = 1
                          
                          row += 1
                          sheet.update_cell(row, 1, Sent_from)
                          sheet.update_cell(row, 2, Mail_Subject)


                          if 'parts' in All_MSG_Body:
                              MSG_parts = All_MSG_Body['parts']
                              print(MSG_parts)
                              for part in MSG_parts:
                                    if part['mimeType'] == 'text/plain':
                                        data = part['body']['data']
                                        decoded_data = base64.urlsafe_b64decode(data).decode('utf-8')
                                    elif part['mimeType'] == 'text/html':
                                        data = part['body']['data']
                                        decoded_txt = base64.urlsafe_b64decode(data).decode('utf-8')
                                        soup = BeautifulSoup(data, 'html.parser')
                                        decoded_data2 = soup.get_text()
                                    else:
                                        decoded_data = ''
                                    print(decoded_data2)


                                    sheet.update_cell(row, 3, decoded_data2)
                                    # sheet.update_cell(row, 4, decoded_data2)

                   
                    
            if  Order_mail == 1 : 
                # print(Order_mail)
                service.users().messages().modify(userId='me', id=message['id'], body={'removeLabelIds': ['UNREAD']}).execute()
                Order_mail = 0
                
                
                          


email_keywords = ["Email", "Outlook", "Gmail", "Mail"]
previous_titles = get_window_titles()
print('Strat....')
while True:
    
    time.sleep(1)
    current_titles = get_window_titles()
    new_windows = [title for title in current_titles if title not in previous_titles]     #----- add the new windows 
    
    for hwnd, title in new_windows:
        if any(keyword.lower() in title.lower() for keyword in email_keywords):
           print('Email_Check') 
           Email_Check()
           exit()
    previous_titles = current_titles