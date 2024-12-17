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
import urllib.parse
import chardet 
from googleapiclient.http import MediaFileUpload
from googleapiclient.discovery import build
import pyautogui
# import pandas as pd


def Google_Sheet_Undate():
    global file_id
    global scope
    global creds_file
    global google_jsonfile
    global drive_service
    global max_retries

    max_retries = 3
    error = 0
    
  


    with open('output.csv', 'r', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        rows = list(reader)

    
    for row in rows:                                                                   # ------  for row in rows----------
        bulletin_number = row[0]  # Get the bulletin number from the first column

        pdf_files = [f for f in os.listdir('download_pdf') if f.startswith(f"{bulletin_number}_")]

        file_links = []
        i = 0
        for pdf_file in pdf_files:                           # -------------  get the links for all files found for each row  ------------
            pdf_path = os.path.join('download_pdf', pdf_file)
            
            # Upload the PDF file to Google Drive
            media_body = MediaFileUpload(pdf_path, resumable=True)
            retries = 0
            while retries < max_retries:
                try:
                    file_drive = drive_service.files().create(body={'name': pdf_file, 'parents': [folder_id]}, media_body=media_body).execute()
                    file_id = file_drive.get('id')
                    i+=1
                    file_link = f'=HYPERLINK("https://drive.google.com/file/d/{file_id}/view","Link{i}")'
                    
                 
                    row.extend([file_link])

                    print(f'uploading {pdf_file}')


                    break
                except Exception as e:
                    print(f"Error uploading file {pdf_file}: {e}")
                    retries += 1
                    if retries < max_retries:
                        print(f"Waiting for 10 seconds before retry #{retries + 1}...")
                        time.sleep(10)
                    else:
                        pyautogui.alert('Error in Uploading file')
                        error = 1
                        return error



    Sheet.insert_rows(rows, row=2)
    row_Num = 2
    Col = 4
    for row in rows :
        if len(row) > 4:
            for Col in range( 5 , len(row)+1):
                Sheet.update_cell(row_Num, Col ,row[Col-1] )
        row_Num += 1

   

                   #------------------------------------------------------------------------------------------------- 







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
      # return error
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
    # return
header = Sheet.row_values(1)
Last_Date = Sheet.row_values(2)[2]    


with open('Url.txt', 'r') as f:
    url = f.read()

if Last_Date == "" :
    Last_Date = '2023-01-01'




#--------------------------------------------------------------------------------------------

os.system('netsh wlan disconnect')
time.sleep(1)

response = requests.get(url)
response.encoding = 'utf-8'
soup = BeautifulSoup(response.content, 'html.parser')

table_rows = soup.find_all('tr')
with open('output.csv', 'w', newline='', encoding='utf-8') as csvfile:
    csv_writer = csv.writer(csvfile)
    for row in table_rows:
        table_data = row.find_all('td')
        if len(table_data) >= 4:
            cell1, cell2, cell3, cell4 = table_data
            Date = cell3.text.strip()

            try:
                Date_and_Time = datetime.strptime(Date, '%d/%m/%Y %H:%M:%S').isoformat()
            except ValueError:
                Date_and_Time = Date

            if Date_and_Time > Last_Date:
                News_No = cell1.text.strip()
                Isin_code = cell2.text.strip()
                News_Body = cell4.text.strip()
                row_data = [News_No, Isin_code, Date_and_Time, News_Body]
                csv_writer.writerow(row_data)

                pdf_links = []
                i = 0
                # Download the PDF file
                pdf_links = cell4.find_all('a', href=lambda href: href)


                for pdf_link in pdf_links:
                    i += 1
                    link = urllib.parse.quote(pdf_link['href'])
                    pdf_url = f'http://online.floor.com/Bulletins/{link}'
                   
                    pdf_filename = f'download_pdf/{News_No}_{i}.pdf'

                    try:
                        response = requests.get(pdf_url)

                        result = chardet.detect(response.content)
                        encoding = result['encoding']
                        response.encoding = encoding
                        with open(pdf_filename, 'wb') as pdf_file:
                            pdf_file.write(response.content)
                        print(f'Downloaded PDF: {pdf_filename}')
                    except Exception as e:
                        print(f'Error downloading PDF: {pdf_filename}')
                        print(f'Error: {e}')

os.system('netsh wlan connect name="Thndr_Brokerage" ssid="Thndr_Brokerage"')

for i in range(40):
    wifi = pywifi.PyWiFi()
    iface = wifi.interfaces()[0]
    connection_status = iface.status()
    time.sleep(1)
    if connection_status == pywifi.const.IFACE_CONNECTED:
        break
if connection_status == 0:
   pyautogui.alert("Wi-Fi Connection is Down")

error = Google_Sheet_Undate()


    
