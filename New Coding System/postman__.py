import sys
import os
import subprocess
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import csv
import pandas as pd
import requests
# import selenium

import autoit  
# import webbrowser  

# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.chrome.options import Options

# import tkinter as tk
from tkinter import messagebox
target_folder = './Users_Data'
CSV_path = os.path.join(target_folder, 'users_data.csv')  
csv_file_path = os.path.normpath(CSV_path) 
                              
# sys.exit()
patch_size = 50


def rename_files(target_folder,csv_file_path):
        
    try: 
        users_data = pd.read_csv(csv_file_path, header=None)  

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
                    pass
                    # print(f'Client name "{client_name}" not found in CSV.')  
    except Exception as e :
        pass
    # sys.exit()


import win32gui
import win32con
from pywinauto import Application  


def get_window_titles():
    def enum_window_titles(hwnd, titles):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            titles.append((hwnd, win32gui.GetWindowText(hwnd)))
        return True

    titles = []
    win32gui.EnumWindows(enum_window_titles, titles)
    return titles



def refresh_chrome():  
    tab_title = 'thndr User'
    titles = get_window_titles()
    chrome_found = False  
    for hwnd, title in titles:  
        if "Google Chrome" in title: 
            chrome_found = True   
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE) 
            win32gui.SetForegroundWindow(hwnd) 
            time.sleep(0.5) 

            app = Application(backend='uia').connect(handle=hwnd)  
            browser = app.window(handle=hwnd)  

            tab_control = browser.child_window(class_name='Chrome_WidgetWin_1')  
            tabs = tab_control.child_window(control_type='Tab', found_index=0) 

                        

            for tab in tabs.ChildWindow():  
                print(tab)
                if tab_title in  tab.window_text():  
                    tab.click_input()  
                    break  

            time.sleep(0.5)  
            from pyautogui import hotkey  
            hotkey('ctrl', 'r')  
            
            break  

    if not chrome_found:  
        print("لا توجد نوافذ Google Chrome مفتوحة.")  

       


def refresh_me():

    # url = "https://usertool.thndr-internal.app"  
    # chrome_path = "C:\Program Files\Google\Chrome\Application\chrome.exe"  

    # subprocess.Popen([chrome_path, url])   

   

    window_title = "thndr User - Google Chrome"

    if autoit.win_exists(window_title):  
        autoit.win_activate(window_title) 
        time.sleep(1)  
        autoit.send("{F5}")  
    else:  
      print('Error')
                    
        

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
            return f"Error with status code {response.status_code}: {response.text}"  
    except Exception as e:
        return f"Exception occurred: {str(e)}"  

   

def download_IDs_CSV (folder_name,target_folder):
    try:
   
        if not os.path.exists(target_folder):  
            os.makedirs(target_folder)  
        command = f"aws s3 cp s3://thndr-coding-prod/{folder_name}/ ./{target_folder}/ --recursive --profile DocumentService-598269941583"
        # command = f"aws s3 sync s3://thndr-coding-prod/{folder_name}/ ./{target_folder}/ --profile DocumentService-598269941583"  

        process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        output, error = process.communicate()
        
        if process.returncode == 0:
            return 'Done'  
    except Exception as e:
       return e

patch = input('enter patch:   ')

refresh_me()
sys.exit()


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
        sheet = client.open_by_url(sheet_url).worksheet('Users Data')
        CSV_sheet = client.open_by_url(sheet_url).worksheet('CSV_File')
else:
    messagebox.showinfo('Error','Can not find Sheet URL')

# Find the last row with values
last_row = len(CSV_sheet.col_values(1)) + 1
    
header =sheet.row_values(1)
ID_index = header.index('New UID')
ID_Number_index = header.index('New ID Number')
google_rows = sheet.get_all_values()



List_of_UIDs = []
patch_Num = 0
for i,row in enumerate(google_rows):  
    if  len(row) > 4 and row[3] and not row[5]:
        
       List_of_UIDs.append(row[3])
       patch_Num+=1
       if patch_Num >= patch_size:
           break
            
print(len(List_of_UIDs))
print(List_of_UIDs)

    
folder_name = Request_Users_IDs(List_of_UIDs)
if isinstance(folder_name, Exception) or folder_name =='None' or folder_name == 'Error':
    # pass
    # messagebox.showinfo("Erore",'ُErore in Request Users IDs function ')
    messagebox.showinfo(folder_name)
    sys.exit

else:
    print(folder_name)
    for i in range(len(List_of_UIDs)*2):

        time.sleep(1)
        print(i)
    feedback =  download_IDs_CSV(folder_name,target_folder)
    if feedback != 'Done':
        messagebox.showinfo("Erore",'ُErore in download IDs CSV function ')
        sys.exit
    else:
        CSV_path = os.path.join(target_folder, 'users_data.csv')  
        try:
            with open(CSV_path , 'r',encoding='utf-8' ) as users_data_file:
                reader = csv.reader(users_data_file)
                rows = list(reader)[1:]
                CSV_sheet.append_rows(rows)  

            
            rename_files(target_folder,csv_file_path)
        except Exception as e:
            messagebox.showinfo('Eror', e)



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







def aaaaaaa():

    data = {
        "id": "06fe197e-c177-4501-ba78-aae47fe9bf01",
        "name": "Thndr Prod",
        "values": [
            {
                "key": "host",
                "value": "https://prod.thndr.app",
                "type": "default",
                "enabled": True
            },
            {
                "key": "admin_token",
                "value": "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJleHAiOjE3Mjk1MDQ4OTA5MjgsImFkbWluX2lkIjoiNzkwNTA1MzMtNjg0MS00ZTRiLWI4YzEtMWRhN2MzMzRkNGNhIiwibmFtZSI6ImZhcmlkIHNoYXdreSIsImVtYWlsIjoiZmFyaWQuc2hhd2t5QHRobmRyLmFwcCJ9.jLC82NuGlimlM_-bGQxElErhEOv4LQJYlaygcQXO4V-o2dWyjjHI3DZ11Sbg2uoEkyiP-H6jnMVdUyg5-N6_mCObz7SOaKr9tNSJcPhME3j8O8RFBapew-nugPRv1PDmzNrWlb4uXxO-lDcrJYAKoYB-FpuLNO_ZsAdrZNtPK4A",
                "type": "default",
                "enabled": True
            }
        ],
        "_postman_variable_scope": "environment",
        "_postman_exported_at": "2024-10-21T08:11:08.568Z",
        "_postman_exported_using": "Postman/11.17.2-241016-0848"
    }

    host = data['values'][0]['value']
    admin_token = data['values'][1]['value']

    url = f"{host}/admin-service/document-service/admin/documents/kyc/bulk-upload-coding-forms"

    headers = {
        "Authorization": admin_token,
        "Content-Type": "application/json"
    }

    data = {
        "users_ids": ['GRmjdGdYHPZNQ4AtVQX0LqNgVM33','cv10yAlDlQh91gllIaHkyXaHPEL2','3dc1Xhkba9TOozYDbsdp7QgXXxQ2']
    }

    response = requests.post(url, headers=headers, json=data)

    if response.status_code == 200:
        print(" JSON:", response.json())
    else:
        print(f"{response.status_code}: {response.text}")






# local_path = "d:\\User_data\\17-10-2024"
# if not os.path.exists(local_path):
#         os.makedirs(local_path)

# run_command("2024-10-14_at_05_0a661ba8-6b8f-4afc-82c0-ac289a736068") 

