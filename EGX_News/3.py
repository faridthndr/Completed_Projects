import requests
from bs4 import BeautifulSoup
import csv
from datetime import datetime
import time
import socket
import os
import pywifi
from pywifi import const
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import csv
from googleapiclient.http import MediaFileUpload
from googleapiclient.discovery import build


def Google_Sheet_Undate():

    folder_id = '1YzDiAOHeTVe6A4NbjvZQJx-JnmYw-NWN'
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_file = 'creds.json'
    google_jsonfile = 'googlekey.json'

    creds = None
    if os.path.exists(creds_file):           # we check if the file path is exist or not
       with open('creds.json', 'r') as f:    # we open the creds file as in read mode and in (f) as a handle of file and load it to creds_json
        creds_json = f.read()
       creds = ServiceAccountCredentials.from_json(creds_json)     # load the creds that will use to access to google sheet from creds_json
    else:
       if not os.path.exists(google_jsonfile):
         
          error = 3
          return error
       else:
          creds = ServiceAccountCredentials.from_json_keyfile_name(google_jsonfile, scope)   # if creds.json file not exist we get the creds by googlekey.json file as usual
          with open(creds_file , 'w') as f:
            f.write(creds.to_json())


    client = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)
    max_retries = 6
  


    if os.path.exists('Sheet_Url.txt'): 
        with open('Sheet_Url.txt', 'r') as u:   #--------------read URL for google sheet
            sheet_url = u.read()

        #Access the Google Sheet
        Sheet = client.open_by_url(sheet_url).worksheet('Daily_News')
    else:
        pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
        return

    header = Sheet.row_values(1)
    Last_Date = Sheet.row_values(2)[2]
    
    

    with open('output.csv', 'r', encoding='utf-8') as csvfile:
               reader = csv.reader(csvfile)
               rows = list(reader)
               for row in rows:  # Skip the header row
                    bulletin_number = row[0]  # Get the bulletin number from the first column
                    pdf_file = os.path.join('download_pdf', f"{bulletin_number}.pdf")

                    if os.path.exists(pdf_file):
                        # Upload the PDF file to Google Drive
                        media_body = MediaFileUpload(pdf_file, resumable=True)     #----- load pdf file to mediafile and let resume = true
                        retries = 0

                        while retries < max_retries:
                            try:
                                file_drive = drive_service.files().create(body={'name': f"{bulletin_number}.pdf", 'parents': [folder_id]}, media_body=media_body).execute()
                                break
                            except Exception as e:
                                print(f"Error uploading file {bulletin_number}.pdf: {e}")
                                retries += 1
                                if retries < max_retries:
                                    print(f"Waiting for 10 seconds before retry #{retries + 1}...")
                                    time.sleep(10)
                                else:
                                    pyautogui.alert('Error in Uploading file')
                                    error = 1
                                    return error
                        # Get the ID of the uploaded file
                        file_id = file_drive.get('id')
                        # Construct the public URL of the uploaded file
                        file_link = f"https://drive.google.com/file/d/{file_id}/view"
                        row.append(file_link)

               Sheet.insert_rows(rows, row=2)







error =  Google_Sheet_Undate()
