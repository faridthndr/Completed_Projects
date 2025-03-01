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
import Main_window as Win                                      # import the file of the window and all interface
import keyboard
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtWidgets, QtGui, QtCore
import sys
app_title = 'EGX News'


global pause 
pause = 0
global drive_service
global folder_id
global file_id



def Google_Sheet_Undate():

    global file_id
    global scope
    global creds_file
    global google_jsonfile
    global drive_service
    global max_retries
    global Sheet
    global folder_id

    max_retries = 3
    error = 0
    

    with open('output.csv', 'r', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        rows = list(reader)

    
    for row in rows:   
        if pause == 2  :
               return 
                # ------  for row in rows----------
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
                    ui.Status_label.setText(f'uploading {pdf_file}')
                    QApplication.processEvents() 


                    break
                except Exception as e:
                    ui.Status_label.setText(f'Error uploading file {pdf_file}: {e}')
                    QApplication.processEvents() 
                    print(f"Error uploading file {pdf_file}: {e}")
                    retries += 1
                    if retries < max_retries:
                        ui.Status_label.setText(f'Waiting for 10 seconds before retry')
                        QApplication.processEvents() 
                        print(f"Waiting for 10 seconds before retry #{retries + 1}...")
                        time.sleep(10)
                    else:
                        ui.Status_label.setText(f'Error in Uploading file')
                        QApplication.processEvents() 
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


def Google_oauth2():


    global file_id
    global scope
    global creds_file
    global google_jsonfile
    global Last_Date
    global Sheet
    global url
    global max_retries
    global Last_Date
    global pause
    global drive_service
    global folder_id

    ui.Status_label.setText("Connecting to Google Sheet ")
    QApplication.processEvents() 

    folder_id = '1YzDiAOHeTVe6A4NbjvZQJx-JnmYw-NWN'
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_file = 'creds.json'
    google_jsonfile = 'googlekey.json'

    os.system('netsh wlan connect name="Thndr_Brokerage" ssid="Thndr_Brokerage"')

    for i in range(40):
        wifi = pywifi.PyWiFi()
        iface = wifi.interfaces()[0]
        connection_status = iface.status()
        time.sleep(1)
        if connection_status == pywifi.const.IFACE_CONNECTED:
            ui.Status_label.setText("Connecting to Wi Fi ")
            QApplication.processEvents() 
            break
    if connection_status == 0:
       ui.Status_label.setText("Wi-Fi Connection is Down")
       QApplication.processEvents()  
       pyautogui.alert("Wi-Fi Connection is Down")
       return

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

    ui.Status_label.setText("Disconnecting Wi Fi ")
    QApplication.processEvents() 

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
                # Last_News_No = int(cell1)
                

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
                            ui.Status_label.setText(f'Downloaded PDF: {News_No}_{i}')
                            QApplication.processEvents() 
                            print(f'Downloaded PDF: {News_No}_{i}')
                        except Exception as e:
                            print(f'Error downloading PDF: {News_No}_{i}')
                            print(f'Error: {e}')

    os.system('netsh wlan connect name="Thndr_Brokerage" ssid="Thndr_Brokerage"')

    for i in range(40):
        wifi = pywifi.PyWiFi()
        iface = wifi.interfaces()[0]
        connection_status = iface.status()
        time.sleep(1)
        if connection_status == pywifi.const.IFACE_CONNECTED:
            ui.Status_label.setText("Connecting to Wi Fi ")
            QApplication.processEvents() 
            break
    if connection_status == 0:
       ui.Status_label.setText("Wi-Fi Connection is Down")
       QApplication.processEvents()  
       pyautogui.alert("Wi-Fi Connection is Down")
       return


def Run_Function():
    global pause
    Timer =  ui.Timer.text()
    while True:

        try:
           Google_oauth2()
           if pause == 2  :
              pause = 0 
              ui.Status_label.setText("Script Puased")
              QApplication.processEvents()    
              return 
        except Exception as e:
          ui.Status_label.setText("Can't Qauth Or getting Url")
          QApplication.processEvents()   
          pyautogui.alert('Error in Google Qauth Or getting Url')   
          return 

        
        try:
            Google_Sheet_Undate()
            ui.Status_label.setText("Waiting for Next Run")
            QApplication.processEvents()   
        except Exception as e:
            return  
        for i in range(int(Timer)*60) :   
            time.sleep(1)
            if pause == 2  :
               ui.Status_label.setText("Script Puased")
               QApplication.processEvents()    
               pause = 0 
               return 
               

def pause_execution():
    global pause
    if pause == 0:
        pause = 1 
        answer = pyautogui.confirm("Do you want to Exit ?", "PAUSE", buttons=["Yes", "No"])
        if answer == "Yes":
            pause = 2
            return
        elif answer == "No":
            pause = 0  
            return
    
    
keyboard.add_hotkey('ALT+ p', pause_execution)


if __name__ == "__main__":
    #import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()

    ui = Win.Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.setWindowTitle(app_title)
    ui.Status_label.setAlignment(QtCore.Qt.AlignCenter)
    ui.start_b.clicked.connect(Run_Function)
    ui.Pause_b.clicked.connect(pause_execution)

    MainWindow.show()
    sys.exit(app.exec_())