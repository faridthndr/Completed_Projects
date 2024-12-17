from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
import os
import gspread
from google.auth.transport.requests import Request


# تحديد الصلاحيات المطلوبة
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.modify']

def main():
    creds = None

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

    results = service.users().messages().list(userId='me', labelIds=['INBOX'], q="is:unread").execute()
    messages = results.get('messages', [])
    print(messages)       

   
    gc = gspread.authorize(creds)
    sheet_url = 'https://docs.google.com/spreadsheets/d/1_CFaVGCKjQK0eWKO56not7LEbjKoNDP9NMAvpJqTyHA/edit?gid=0#gid=0'
    sheet = gc.open_by_url(sheet_url).worksheet('Emails')

    sheet.update_cell(1, 1, "Updated Value")

if __name__ == '__main__':
    main()
