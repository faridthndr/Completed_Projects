import sys
import os
import subprocess
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import csv
import socket 
import pandas as pd
import requests
import win32gui
import autoit
import threading
from PyQt5.QtWidgets import QApplication ,QMessageBox
from PyQt5 import QtWidgets  
import keyboard
import Main_window as window
import random  
from pynput.mouse import  Controller  


# import tkinter as tk  
from tkinter import messagebox

APPTitle = "New Coding System"
google_rows = None
EGX_Code_sheet = None
CSV_sheet = None
pause_flag = False
Users_Data_Sheet = None
autoit_script = './Data/Download_Codes_TCP.au3'
target_folder = './Users_Data'
CSV_path = os.path.join(target_folder, 'users_data.csv')  
EGX_COde_file_path = 'New folder'  # قم بتغيير هذا المسار  

users_data_path = os.path.normpath(CSV_path) 
All_Users_data_path = os.path.join(target_folder,'All_Users_data.csv')

# sys.exit()
Batch_size = 50


def pause_execution():
    global pause_flag,run
    if pause_flag == False  :        
        answer = messagebox.askquestion("Confirm Action", "Do you want to exit?")
        if answer == "yes":
            pause_flag = True
            win.title.setText("Operations have been cancelled..")
            # win.start.clicked.connect(test)
            QApplication.processEvents()
            return pause_flag
        elif answer == "no":
            pause_flag = False
            return pause_flag



# def Autoit_TCP_Call():
        
#     host = '127.0.0.1'   
#     port = 65432        

#     with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:  
#         s.bind((host, port))  
#         s.listen()  
#         print(f"Listening on {host}:{port}")  
#         subprocess.Popen(['./Data/autoit3',autoit_script], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
#         while True:  
            
            
#             conn, addr = s.accept()  
#             with conn:  
#                 print(f"Connected by {addr}") 
#                 data = conn.recv(1024) 
#                 if not data:  
#                     break  
#                 message = data.decode('utf-8')  
#                 print(f"Received data: {message}")

#                 return data
     


# def human_move():  
#     mouse = Controller() 
#     duration = 0.4
#     for i in range (15):
#         if pause_flag :
#             return
#         destination = (random.randint(200, 1400 - 1), random.randint(100, 700 - 1))
#         start_position = mouse.position  
#         distance_x = destination[0] - start_position[0]  
#         distance_y = destination[1] - start_position[1]  

#         steps = 100  
#         pause_duration = duration / steps  

#         for i in range(steps):  
#             x = int(start_position[0] + (distance_x * (i / steps)) + random.uniform(-1, 1))  
#             y = int(start_position[1] + (distance_y * (i / steps)) + random.uniform(-1, 1))  
#             mouse.position = (x, y)  
#             time.sleep(pause_duration)  




keyboard.add_hotkey('ALT+ p', pause_execution)



def rename_files(target_folder,Users_data_path):
        
    try: 
        users_data = pd.read_csv(Users_data_path, header=None)  

        users_dict = dict(zip(users_data[0], users_data[9]))

        for filename in os.listdir(target_folder):  
            if "_front_id_" in filename or "_back_id_" in filename:  

                parts = filename.split("_")  
                # client_name = " ".join(parts[2:-1])  
                client_name = " ".join(parts[3:]).replace(".jpg", "").replace(".png", "")  
                # print(client_name)
                if client_name in users_dict:  
                    card_number = users_dict[client_name]  
                    new_filename = f"{card_number}_{parts[1]}_{parts[2]}.jpg"  
                    os.rename(os.path.join(target_folder, filename), os.path.join(target_folder, new_filename))  
                    print(f'Renamed: {filename} to {new_filename}')  
                    
                else:  
                    # return False
                    messagebox.showinfo("Error","Can not find some filename in CSV file")

        return True
    except Exception as e :
       messagebox.showinfo("Error",e)
       return False
    # sys.exit()



# def Calling_Soket():
    
#     host = '127.0.0.1'   
#     port = 65432         

#     with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:  
#         s.bind((host, port))  
#         s.listen()  
#         print(f"Listening on {host}:{port}")  
        
#         while True: 
#             try: 
#                 conn, addr = s.accept()  
#                 with conn:  
#                     print(f"Connected by {addr}")  
#                     data = conn.recv(1024) 
#                     if not data:  
#                         break  
#                     message = data.decode('utf-8')
#                     message =  message.split(',')  
#                     # value_to_compare = message[0]
#                     # if  value_to_compare == "" :
#                     return True


#             except Exception as e:
#                 return False
#                 # conn.close()    # if we need to close the connection


def read_excel_files(folder_path):  
    for filename in os.listdir(folder_path):  
        if filename.endswith('.xlsx'):  
            file_path = os.path.join(folder_path, filename)  
            data = pd.read_excel(file_path)  
            for column in data.select_dtypes(include=['datetime', 'datetimetz']).columns:
                data[column] = data[column].apply(lambda x: x.isoformat() if pd.notnull(x) else None)

            # data = data.astype(str)
            data_list = data.values.tolist()
            return data_list
           


def sync_data(users_data_path,all_users_data_path):
    
    if  os.path.exists(users_data_path):  
       
        users_data = pd.read_csv(users_data_path, header=None)  
       

        if not os.path.exists(all_users_data_path):  
            all_users_data = pd.DataFrame(columns=users_data.columns) 
            all_users_data.to_csv(all_users_data_path, header=False, index=False)  
        else:  
            all_users_data = pd.read_csv(all_users_data_path, header=None)  

        existing_users_column = all_users_data[8].tolist() 

        for index, row in users_data.iterrows():  
            if row[8] not in existing_users_column:  
                all_users_data = all_users_data.append(row, ignore_index=True)  

        all_users_data.to_csv(all_users_data_path, header=None, index=False)



# def get_window_titles():
#     def enum_window_titles(hwnd, titles):
#         if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
#             titles.append((hwnd, win32gui.GetWindowText(hwnd)))
#         return True

#     titles = []
#     win32gui.EnumWindows(enum_window_titles, titles)
#     return titles



def win_activate(title,time_out):
    try:
        autoit.win_activate(title)
        autoit.win_wait_active(title,time_out)
        return True 

    except Exception as e:
        messagebox.showinfo("Can not find Window","Can not find Window")
        return False



def refresh_Chrome():

    # url = "https://usertool.thndr-internal.app"  
    # chrome_path = "C:\Program Files\Google\Chrome\Application\chrome.exe"  

    # subprocess.Popen([chrome_path, url])   

   

    window_title = "thndr User - Google Chrome"

    if autoit.win_exists(window_title):  
        autoit.win_activate(window_title) 
        time.sleep(1)  
        autoit.send("{F5}")
        return True  
    else:  
       messagebox.showinfo("Error",'Cannot find the Chrome window to activate. Please ensure that the user tool window is open ' )
       return False
                    
        

def Request_Users_IDs(Users_UIDs):

    try:
        url = "https://prod.thndr.app/admin-service/document-service/admin/documents/kyc/bulk-upload-coding-forms"
        with  open('Token.txt', 'r') as r: 
            token = r.read()

        headers = {
            "Authorization":f"Bearer {token}",
            "Content-Type": "application/json"
        }

        data = {
            "users_ids": Users_UIDs   
        }

        response = requests.post(url, headers=headers, json=data)

        if response.status_code == 200:
            response_data = response.json()
            folder_name = response_data['folder_name']
            return folder_name
        else:  
            return False  

    except Exception as e:
        return False 

  

def download_IDs_CSV (folder_name,target_folder):
    try:
   
        if not os.path.exists(target_folder):  
            os.makedirs(target_folder)  
        command = f"aws s3 cp s3://thndr-coding-prod/{folder_name}/ ./{target_folder}/ --recursive --profile DocumentService-598269941583"
        # command = f"aws s3 sync s3://thndr-coding-prod/{folder_name}/ ./{target_folder}/ --profile DocumentService-598269941583"  

        process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        output, error = process.communicate()
        
        if process.returncode == 0:
            return True
    except Exception as e:
       return False



def google_Auth():

    global sheet, google_rows,CSV_sheet,EGX_Code_sheet,Users_Data_Sheet
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_file = 'creds.json'
        google_jsonfile = 'coding-new-users.json'

        creds = None
        if os.path.exists(creds_file):         
            with open('creds.json', 'r') as f:  
                creds_json = f.read()
                creds = ServiceAccountCredentials.from_json(creds_json)     # load the creds that will use to access to google sheet from creds_json
        else:
            if not os.path.exists(google_jsonfile):     
                error = 3
            else:
                creds = ServiceAccountCredentials.from_json_keyfile_name(google_jsonfile, scope)   # if creds.json file not exist we get the creds by googlekey.json file as usual
                with open(creds_file , 'w') as f:
                    f.write(creds.to_json())

        client = gspread.authorize(creds)


        if os.path.exists('Sheet_Url.txt'): 
                with open('Sheet_Url.txt', 'r') as u:   #--------------read URL for google sheet
                    sheet_url = u.read()

                #Access the Google Sheet
                Users_Data_Sheet = client.open_by_url(sheet_url).worksheet('Users Data')
                EGX_Code_sheet = client.open_by_url(sheet_url).worksheet('EGX_Codes')
                CSV_sheet = client.open_by_url(sheet_url).worksheet('CSV_File')
        else:
            messagebox.showinfo('Error','Can not find Sheet URL')
            return False
        # Find the last row with values
        last_row = len(CSV_sheet.col_values(1)) + 1
            
        # header =sheet.row_values(1)
        # ID_index = header.index('New UID')
        # ID_Number_index = header.index('New ID Number')
        # google_rows = sheet.get_all_values()

        

        return True
    except Exception as e:
        messagebox.showinfo("Error",{e})
        return False



def start():
    
    global sheet,google_rows,CSV_sheet,target_folder,users_data_path,pause_flag,MCDR_title,EGX_COde_file_path,EGX_Code_sheet,Users_Data_Sheet

    pause_flag = False
    win.start.clicked.disconnect()
    win.start.setEnabled(False)
    Timer = int (win.timer.text())
    Last_Refresh = time.time()

    while True:
        
        try:

            google_rows = Users_Data_Sheet.get_all_values()
            
            if not win.Batch.text().isdigit() or win.Batch.text() == "":     
                messagebox.showinfo("Error", 'Some required options are not filled in or there are empty fields in the list. Please complete all necessary selections.')
                return

            if pause_flag:
                pause_flag = False
                return

            Batch_size = int(win.Batch.text())
            if win.refreshchrome.isChecked():    # ---------------- check for refresh -------------------
                if not refresh_Chrome():
                    pass
                    # return
                if not  win_activate(APPTitle,2):
                     win.title.setText('Can not find Window ')
                     QApplication.processEvents()
                    #  return
                win.title.setText('Refresh the User Tool screen')
                QApplication.processEvents()

            if win.usersIDs.isChecked():                # -------------------- request Users IDs Photo -------------
                List_of_UIDs = []
                patch_Num = 0
                for i, row in enumerate(google_rows):
                    if len(row) > 4 and row[3] and not row[5]:
                        List_of_UIDs.append(row[3])
                        patch_Num += 1
                        if patch_Num >= Batch_size:
                            break
                    if pause_flag:
                        pause_flag = False
                        return

                print(List_of_UIDs)
                
                folder_name = Request_Users_IDs(List_of_UIDs)
                if not folder_name:
                    messagebox.showinfo("Error", 'Error in Request Users IDs function')
                    return

                else:
                    print(folder_name)
                    for i in range(len(List_of_UIDs) * 2 + 10  + 1):
                        time.sleep(1)
                        win.title.setText(f'Time remaining: {len(List_of_UIDs) * 2 + 10 - i} seconds')
                        QApplication.processEvents()
                        if pause_flag:
                            pause_flag = False
                            return
                        current_time = time.time()
                        if win.refreshchrome.isChecked() and (current_time -Last_Refresh) > 120:    #--------- check for time to refresh user tool screen ----
                            #    win_activate(MCDR_title,3)
                               if not refresh_Chrome():
                                    pass
                               if not  win_activate(APPTitle,2):
                                    win.title.setText('Can not find Window ')
                                    QApplication.processEvents()
                                    return
                               win.title.setText('Refresh the User Tool screen')
                               QApplication.processEvents()
                               Last_Refresh = time.time()


                    win.title.setText('Downloading OF IDs Photos.')
                    QApplication.processEvents()
                    if not download_IDs_CSV(folder_name, target_folder):
                        messagebox.showinfo("Error", 'Error in download IDs CSV function')
                        return
                    else:
                        CSV_path = os.path.join(target_folder, 'users_data.csv')
                        try:
                            with open(CSV_path, 'r', encoding='utf-8') as users_data_file:
                                reader = csv.reader(users_data_file)
                                rows = list(reader)[1:]
                                CSV_sheet.append_rows(rows)

                        except Exception as e:
                            messagebox.showinfo("Error", 'Error while attempting to write data to Google Sheets. The process will stop now. Please try again.')
                            return

                        if not rename_files(target_folder, users_data_path):
                            # messagebox.showinfo("Error", 'An error occurred while trying to rename the files in the folder.')
                            return

                        win.title.setText('Photo files have been successfully renamed')
                        QApplication.processEvents()

            if win.UsersCodes.isChecked():         #---------- if we Select update users Codes ----------- 
               
               data = read_excel_files('New folder')
               EGX_Code_sheet.batch_clear(['A2:I'])
               EGX_Code_sheet.update(values= data, range_name='A2:I')

            
            for i in range(Timer):            #---------------- loop for delay time --------------------------
                current_time = time.time()
                if win.refreshchrome.isChecked() and (current_time -Last_Refresh) > 120:    #--------- check for time to refresh user tool screen ----
                    #    win_activate(MCDR_title,3)
                        if not refresh_Chrome():
                            pass
                            # return
                        if not  win_activate(APPTitle,2):
                            messagebox.showinfo("Error","Can not find Window ")
                            return
                        win.title.setText('Refresh the User Tool screen')
                        QApplication.processEvents()
                        Last_Refresh = time.time()

                win.title.setText(f'Waiting fot next Run: {Timer - i} seconds')
                QApplication.processEvents()
                time.sleep(1)    
                if pause_flag:
                    pause_flag = False
                    return

            

        finally:
            win.start.setEnabled(True)  
            win.start.clicked.connect(start) 

    current_time = time.time()
    # if win.UsersCodes.isChecked() and (current_time -Last_Refresh) > 20:                                            #-------------------------------
    #         # current_time = time.time()
    #         # if(current_time -Last_Refresh) > 20 :
    #         win_activate(MCDR_title,3)
    #         Run_human_move.start()    
                


if __name__ == "__main__" :
    
    app = QtWidgets.QApplication(sys.argv)
    MainWin = QtWidgets.QMainWindow()
    win = window.Ui_MainWin()
    win.setupUi(MainWin)
    MainWin.setWindowTitle(APPTitle)
    win.refreshchrome.setChecked(True)
    win.usersIDs.setChecked(True)
    win.UsersCodes.setChecked(True)
    win.timer.setText("300")
    win.Batch.setText("50")
    win.title.setText('Starting Google authentication.')
    
    MainWin.show()
    
    QApplication.processEvents()
    if not  google_Auth():
       messagebox.showinfo("Error","Unable to connect to Google Sheets. Please check the authentication files, the URL file, and the internet connections. ")
       sys.exit()
    win.title.setText('Ready to Start')
    QApplication.processEvents()
    win.start.clicked.connect(start)
    win.pause.clicked.connect(pause_execution)
    

    sys.exit(app.exec_())
    


# Batch_size = int(input('enter patch:   '))



# refresh_me()
# exit()

# List_of_UIDs = []
# patch_Num = 0
# for i,row in enumerate(google_rows):  
#     if  len(row) > 4 and row[3] and not row[5]:
        
#        List_of_UIDs.append(row[3])
#        patch_Num+=1
#        if patch_Num >= Batch_size:
#            break
            
# print(len(List_of_UIDs))
# print(List_of_UIDs)

    
# folder_name = Request_Users_IDs(List_of_UIDs)
# if isinstance(folder_name, Exception) or folder_name =='None' or folder_name == 'Error':
#     # pass
#     # messagebox.showinfo("Erore",'ُErore in Request Users IDs function ')
#     messagebox.showinfo(folder_name)
#     exit()

# else:
#     print(folder_name)
#     for i in range(len(List_of_UIDs)*2):

#         time.sleep(1)
#         print(i)
#     feedback =  download_IDs_CSV(folder_name,target_folder)
#     if feedback != 'Done':
#         messagebox.showinfo("Erore",'ُErore in download IDs CSV function ')
#         exit()
#     else:
#         CSV_path = os.path.join(target_folder, 'users_data.csv')  
#         try:
#             with open(CSV_path , 'r',encoding='utf-8' ) as users_data_file:
#                 reader = csv.reader(users_data_file)
#                 rows = list(reader)[1:]
#                 CSV_sheet.append_rows(rows)  

            
#             rename_files(target_folder,users_data_path)
#         except Exception as e:
#             messagebox.showinfo('Eror', e)
       



# def upload_user_ids(url, user_ids):

#     url = "{{host}}/admin-service/document-service/admin/documents/kyc/bulk-upload-coding-forms"

#     headers = {
#         'Authorization': 'Bearer your_token' 
#     }
#     data = {
#         "users_ids": user_ids
#     }

#     response = requests.post(url, headers=headers, json=data)

#     if response.status_code == 200:
#         response_data = response.json()
#         return response_data['folder_name']   

#     else:
#         print(" ", response.status_code)
#         return None







# local_path = "d:\\User_data\\17-10-2024"
# if not os.path.exists(local_path):
#         os.makedirs(local_path)

# run_command("2024-10-14_at_05_0a661ba8-6b8f-4afc-82c0-ac289a736068") 

