import pyodbc
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import csv
import codecs  # Import the codecs module for encoding support
#from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtWidgets, QtGui, QtCore
import ctypes

import datetime
from decimal import Decimal
import pyautogui
import time
import mainwindow as Win                                      # import the file of the window and all interface
from pywinauto import Application
import pyperclip
import keyboard
import threading
import pandas as pd
#import pytesseract
import requests
import os
import json
import sys
import tkinter as tk
from tkinter import messagebox

from win32con import WM_INPUTLANGCHANGEREQUEST
from win32gui import GetForegroundWindow
from win32api import SendMessage
#import autoit
from pynput.keyboard import Controller, Key 
keyboard2 = Controller()



T0 ="BRP_800                                       "
T1 ="BRP_702                                       "
T2 ="BRP_410                                       "
Book_Cancel ="BRP_951                                       "

Last_Alert__Run = 0
Timer = 0
pause = 0
LastUpdateDate = '2024-01-01 10:00:00.973000'
issue_flag = 0
app_title = 'Session Support'
SymbolCode = ''
QTY = 0
MCDR_Code = 0
CustodianCode = ''
UnifiedCode = ''
comment = ''
Stock_Num = '0'
cursor = None
Alert_Sheet = None


#pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Thndr\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
pyautogui.PAUSE = 0.1

path = os.getcwd() + '\\data\\'


def Change_language():
    time.sleep(0.2)
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    curr_window = user32.GetForegroundWindow()
    thread_id = user32.GetWindowThreadProcessId(curr_window, 0)
    klid = user32.GetKeyboardLayout(thread_id)
    lid = klid & (2**16 - 1)
    lid_hex = hex(lid)
    if lid_hex == '0xc01':
        #if SendMessage( GetForegroundWindow(), WM_INPUTLANGCHANGEREQUEST, 0, 0x4090409) == 0:
        SendMessage( GetForegroundWindow(), WM_INPUTLANGCHANGEREQUEST, 0, 0x4090409) 
       



stop_flag = threading.Event()

current_date = current_date = datetime.datetime.now()
new_date = current_date + datetime.timedelta(days=30)
Expirydate =new_date.strftime("%d\\%m\\%Y")



def WinActivate(app_title, time_out):
    global pause
    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return

        
    try:
        app = Application(backend='uia').connect(title=app_title,timeout=time_out,visible_only=False)
        window = app.window(title=app_title)
        window.set_focus()
        return True
    except Exception as e:
        #print(f"App window with title '{app_title}' not found or could not be activated.")
        #time.sleep(1)
        return False


def detect_and_click():

    global issue_flag

    while not stop_flag.is_set()  :
        try:
            # Check if the image appears on the screen
            if pyautogui.locateOnScreen(path+'Issue_Num.png') is not None :
                issue_flag = 1
                time.sleep(0.3)
                # Perform the click action on the image
                pyautogui.click(path+'Issue_Num_Ok.png')
                issue_flag = 0
                #break
        except Exception as f:
            messagebox.showinfo("Error",f) 
            time.sleep(1)       

image_detection_thread = threading.Thread(target=detect_and_click)
image_detection_thread.start()



def clickOn(img,offset=0,click=1,timeout = 0):
    global pause
    global path

    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return


    for _ in range(timeout):        
        try:
            x,y = pyautogui.locateCenterOnScreen(path+img,confidence = 0.8)                  # 90% matching the Img
            
            
        except Exception as e:
            time.sleep(0.1)
            
        else:
            if click == 2 :
               pyautogui.doubleClick(x- offset, y)

            elif click == 1:
                pyautogui.leftClick(x-offset,y, interval=0.001)

            return x , y
            break
    print("Can't Find "+ img )
 

def Detect_img(img1,img2,img3, time_out):
    global path
    for _ in range(time_out):

        
        try:
           x,y =  pyautogui.locateCenterOnScreen(path+img1, confidence=0.8)
        except Exception as e:
          time.sleep(0.1)
          x, y = None, None
          img = None
        else:
          img = 1
          return x, y , img
          break

        if img2 != 0 :

            try:
                x,y =  pyautogui.locateCenterOnScreen(path+img2, confidence=0.8)
            except Exception as e:
                time.sleep(0.1)
                x, y = None, None
                img = None
            else:
                img = 2
                return x, y , img
                break
        if img3 != 0 : 
            try:
                x,y =  pyautogui.locateCenterOnScreen(path+img3, confidence=0.8)
            except Exception as e:
                time.sleep(0.1)
                x, y = None, None
                img = None
            else:
                img = 3
                return x, y , img
                break
        

    x, y = None, None
    img = None
    print("Can't Find "+ img1 )

    return x ,y,img



def T0_call(MCDR_Code,UNI_Code,sheet,i,Action_index):

    global pause,T0,issue_flag,SymbolCode,comment
    issue_flag = 0
    

    while  pause ==1:
        if pause == 0 :
            break
        elif pause == 2 :
            return
    try:
     
     if not WinActivate(T0,3):
        messagebox.showinfo("Error","Can not find MCDR Window")
        return False

     Change_language() 
     time.sleep(0.1)
     clickOn("client.png",90,2,5)
     time.sleep(0.1)
     pyautogui.typewrite(UNI_Code)
     clickOn("stock.png",90,2,5)
     pyautogui.typewrite(MCDR_Code)
     time.sleep(0.2)
     clickOn("search.png",0,1,5)
     time.sleep(0.1)
     if issue_flag == 1:
         time.sleep(1.5)
         issue_flag = 0

     #clickOn("Issue_Num_Ok.png",0,1,5)     #----------------------------------------
     clickOn("search.png",0,1,10)
     time.sleep(0.3)


     clickOn('record.png', 0, 1,20)
     x,y,img = Detect_img("confirm_Win.png","No_Qty_Available.png",0,30)

     if img == 1:
         clickOn('The_Qty_toBook.png',95,2,10)
         time.sleep(0.15)
         pyperclip.copy('')
           # Simulate pressing Ctrl + C to copy the text
         pyautogui.hotkey('ctrl', 'c')
         time.sleep(0.3)
         pyautogui.hotkey('ctrl', 'c')
         time.sleep(0.3)
         Stock_Num = pyperclip.paste() 

         if Stock_Num.strip().isdigit() and int(Stock_Num) > 0:

             time.sleep(0.1)
             clickOn("booked_Button.png", 0, 1,10)
 
             x,y,img = Detect_img("Max_QTY.png","Done_MSG.png",0,20)
             if img == 1 :
                 time.sleep(0.3)
                 clickOn("done.png",0,1,30) 
                 time.sleep(0.3)
                 clickOn("back.png",0,1,30) 
                 comment = 'MAx QTY'   
                #  sheet.update_cell(i+1, Action_index+1, 'MAx QTY')
                 return True
             elif img == 2:
                 time.sleep(0.3)
                 clickOn("done.png", 0, 1,10)
                 comment = Stock_Num + '  Booked'
                #  sheet.update_cell(i+1, Action_index+1, Stock_Num + '  Booked')
                 return True

         else: 
               clickOn('back.png',0,1,15)
               time.sleep(0.15)   
               comment = 'MAx QTY' 
            #    sheet.update_cell(i+1, Action_index+1, 'MAx_QTY')
               return True
     if img ==2:
             clickOn("done.png",0,1,30)
             # Stock_Num = 'Sold or No Available Qty' 
             comment = 'Sold or No Available Qty'
            #  sheet.update_cell(i+1, Action_index+1, 'Sold or No Available Qty')
             return  True
    except Exception as e:
        pyautogui.alert(e,"Erro")
        ui.Status_label.setText(e)
        QApplication.processEvents()
        return False
 

def T1_call(MCDR_Code,UNI_Code,sheet,i,Action_index):
    global pause,issue_flag,SymbolCode,comment
    issue_flag = 0
    try:

        while  True:
            if pause == 0 :
                break
            elif pause == 2 :
                return


        if not WinActivate(T1,3):
            messagebox.showinfo("Error","Can not find MCDR Window")
            return False

        Change_language() 
        time.sleep(0.1)
        clickOn("client.png",90,2,20)
        time.sleep(0.1)
        pyautogui.typewrite(UNI_Code)
        time.sleep(0.1)
        clickOn("The_Stock.png",90,2,20)
        pyautogui.typewrite(MCDR_Code)
        time.sleep(0.2)
        clickOn("search.png",0,1,20)
        
        time.sleep(0.4)

        if issue_flag == 1:
            time.sleep(0.7)
            # issue_flag = 0
            Detect_img('Issue_Num.png',0,0,20)
            # time.sleep(.6)
        if issue_flag == 1:
            time.sleep(0.6)
            # issue_flag = 0

        time.sleep(0.4)    
        # clickOn('No_Issue_Select.png',0,1,20)  
        clickOn("search.png",0,1,20)

        time.sleep(0.3)
        clickOn('Boked.png', 0, 1,20)
        time.sleep(0.2)
        x,y,img = Detect_img('record_sell_order.png',0,0,15)
        if img == None:
            Stock_Num = None
            return False 

        pyperclip.copy('')
        # Simulate pressing Ctrl + C to copy the text
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.3)
        # Retrieve the copied text from the clipboard
        Stock_Num = pyperclip.paste()
        if Stock_Num.strip().isdigit() and int(Stock_Num) > 0:

            clickOn('save_button.png', 0, 1,20)
            time.sleep(0.2)
            clickOn('done.png',0,1,20)

            number = 0  
            number = int(Stock_Num)
            if number > 0  :
                comment = comment +  Stock_Num + ' ' + 'Booked'
                #  sheet.update_cell(i+1, Action_index+1, Stock_Num + ' ' + 'Booked')
                return True   


        else:
            clickOn('back.png',0,1,20) 
            comment = comment +  'No Stock'
            # sheet.update_cell(i+1, Action_index+1, 'No Stock')
            return True

        # return Stock_Num
    except Exception as e:
        pyautogui.alert(e,"Erro")
        ui.Status_label.setText(e)
        QApplication.processEvents()
        return False



def Un_Booking():
    global pause,UnifiedCode,Symbol_Num,Stock_Num,CustodianCode,QTY,comment

    
    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return

    if not WinActivate(Book_Cancel,3):
       messagebox.showinfo("Error","Can not find MCDR Window")
       Stock_Num = 0 
       return False
    else:
       time.sleep(0.2)
       clickOn('Custodian_Code.png',70,2,10)
       time.sleep(0.2)
       keyboard2.type(CustodianCode)

       time.sleep(0.3)
       keyboard2.type('\t')  
       time.sleep(0.2)
       clickOn("The_Client2.png",90,2,30)
       Change_language() 

       time.sleep(0.5)
       keyboard2.type(UnifiedCode)
      
       keyboard2.type('\t')  

       time.sleep(0.5)
       clickOn("The_Stock.png",90,2,30)
       time.sleep(0.4)
       keyboard2.type(MCDR_Code)
       time.sleep(0.4)
       keyboard2.type('\t')  
       time.sleep(0.5)               
       keyboard2.type('31122000')
       time.sleep(0.5)
       clickOn('Search.png',0,1,30)
       time.sleep(2)
       
       try:
            x,y,img = Detect_img("Remain.png",0,0,30)
       except Exception as e:
            pass
       else:
            pyautogui.doubleClick(x, y+35)
            time.sleep(0.2)            
            pyperclip.copy('')
            time.sleep(0.4)

            # Press CTRL + C
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'c')

            time.sleep(0.4)
            Stock_Num = pyperclip.paste()
            
            
            if Stock_Num.strip().isdigit() and int(Stock_Num) > 0 :
               
                clickOn('Delete.png',0,1,40)
                time.sleep(0.2)
                x,y,img = Detect_img('Order_Is_Deleted.png','Cannot_delete.png','Delete_Comment.png',30)
                if img == 1:
                    time.sleep(0.3)
                    clickOn('Order_Is_Deleted_ok.png',0,1,20)
                    time.sleep(0.2)
                    if int(Stock_Num) > int(QTY):
                        QTY = (Stock_Num)
                        print(QTY)
                    comment = Stock_Num + ' Unbooked '
                    return Stock_Num
                elif img == 2:
                    time.sleep(0.2)
                    clickOn('FRM_40401_Ok.png',0,1,8)
                    time.sleep(0.4)
                    Stock_Num = 0
                    # error = 3
                    comment = ' Qty in MkT '
                    return Stock_Num
                elif img == 3 :
                    time.sleep(0.2)    
                    clickOn('Delete_Comment.png',90,2,10)
                    keyboard2.type('cancel')
                    time.sleep(0.2)
                    clickOn('OK.png',0,1,30)
                    time.sleep(0.4)
                    x,y,img = Detect_img('Cannot_delete.png','Delete_Done.png','FRM_40401.png',30)
                    if img == 1:
                        time.sleep(0.2)
                        clickOn('Cannot_delete_ok.png',0,1,8)
                        time.sleep(0.4)
                        clickOn('back2.png',0,1,10)
                        time.sleep(0.4)
                        Stock_Num = 0
                        # error = 3
                        comment = ' Qty in MkT '
                        return Stock_Num

                    elif img == 2:
                        time.sleep(0.3)
                        clickOn('Delete_Done_ok.png',0,1,40)
                        time.sleep(0.2)
                        # error = 0
                        if int(Stock_Num) > int(QTY):
                            QTY = (Stock_Num)
                            print(QTY)
                        comment =  Stock_Num + ' Unbooked '
                        return Stock_Num
                    elif img == 3 :
                        time.sleep(0.3)
                        clickOn('FRM_40401_Ok.png',0,1,10) 
                        time.sleep(0.3)
                        x,y,img = Detect_img('Order_Is_Deleted.png','Delete_Done.png',0,30)
                        if  img == 1:
                            time.sleep(0.35)
                            clickOn('Order_Is_Deleted_ok.png',0,1,10)
                            time.sleep(0.2)
                            # error = 0
                            if int(Stock_Num) > int(QTY):
                                QTY = (Stock_Num)
                                print(QTY)
                            comment = Stock_Num + ' Unbooked '
                            return Stock_Num
                        elif img == 2:
                            time.sleep(0.35)
                            clickOn('Delete_Done_Ok.png',0,1,10)
                            time.sleep(0.2)
                            if int(Stock_Num) > int(QTY):
                                QTY = (Stock_Num)
                                print(QTY)
                            # error = 0
                            comment = Stock_Num + ' Unbooked '
                            return Stock_Num
                        else:

                            # error = 3
                            Stock_Num = 0
                            if int(Stock_Num) > int(QTY):
                                QTY = (Stock_Num)
                                print(QTY)
                            comment = Stock_Num + ' Unbooked '
                            return Stock_Num
                            
                if img == None:
                        Stock_Num  = 0
                        # error = 3
                        comment = 'Qty in MkT'
                        return Stock_Num
            else:
                # error = 2
                Stock_Num = 0
                if int(Stock_Num) > int(QTY):
                        QTY = (Stock_Num)
                        print(QTY)
                comment = str(Stock_Num) + ' Unbooked '
                return Stock_Num



def Booking_T2():

    global Valed_Date
    global T2
    global SymbolCode
    global QTY
    global CustodianCode
    global UnifiedCode
    global Expirydate
    global comment
    global Stock_Num

    try:
        while True:
            if pause == 0 :
                break
            elif pause == 2 :
                return
  
        if not WinActivate(T2,4):
            messagebox.showinfo("Error","Can not find MCDR Window")
            return False
        else:
            time.sleep(0.2)

            clickOn('SymbolCode.png',120,2,30)

            #Change_language() 

            keyboard2.type(SymbolCode)
            clickOn('The_QTY.png',120,2,30)
            #if QTY < Stock_Num :
            #    QTY = Stock_Num
            keyboard2.type(QTY)
            time.sleep(0.3)
            clickOn('CustodianCode.png',120,2,30)
            keyboard2.type(CustodianCode)
            time.sleep(0.3)
            clickOn('The_Client.png',120,2,30)
            keyboard2.type(UnifiedCode)
            time.sleep(0.3)
            pyautogui.typewrite('\t') 
            time.sleep(0.3)
        
            keyboard2.type(Expirydate)
            pyautogui.typewrite('\t') 
            clickOn('Save.png',0,1,30)
            
            x,y,img = Detect_img ('Not_Enough_Balance.png','Qty_Is_Booked.png','not_listed.png',40)
            if img == 1: 
                clickOn('Delete_Done_Ok.png',0,1,20)
                omment = comment + ' - No Balance for ' + QTY
                return  True
            elif img == 2 :
                time.sleep(0.2)
                clickOn('Qty_Is_Booked_Ok.png',0,1,40)
                comment = comment  + QTY + ' Booked'
                return True
            elif img == 3 :
                time.sleep(0.2)
                clickOn('Cannot_delete_ok.png',0,1,10)
                time.sleep(0.2)
                comment = comment + ' Code Suspended  '
                return True
    except Exception as e:
        messagebox.showinfo("Error", e)
        return False


def send_slack_notification(query_results):  
    Slack_url = 'https://hooks.slack.com/services/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' 
    # Slack_url =  'https://hooks.slack.com/services/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'   
    csv_data = []   
    csv_data.append("    Timestamp  ,         Message     ,    Count  ")  

    for row in query_results:  
        timestamp = row[0]
        message = row[1].strip()  
        count = row[2]  
        
        csv_data.append(f"{timestamp},     {message}    ,  {count}  ")  

    csv_content = "\n\n".join(csv_data)    
    
    payload = {  
        "text": f"Warning: Number of rejection messages in the last inquiry is:\n```\n{csv_content}\n```"  
    }  
    
    headers = {  
        'Content-Type': 'application/json'  
    }  
    
    response = requests.post(Slack_url, json=payload, headers=headers)  

    if response.status_code != 200:  
        return False 
    else:
        return True



def  MCDR_ALert():
    global Timer,stop_flag,cursor,Alert_Sheet,Last_Alert__Run

    
    if cursor != None and Alert_Sheet != None:      
        try:
            with open ('Alert_Query.txt','r') as m:
                    Alert_Query  = m.read()
                    Alert_Query = Alert_Query.replace('--*',f'-{Timer}' )
            
            cursor.execute(Alert_Query)
            alert_results = cursor.fetchall()
            alert_results = [list(row) for row in alert_results] 
            for row in alert_results:
                row[0] = row[0].strftime('%Y-%m-%d %H:%M:%S')

            Alert_header = Alert_Sheet.row_values (1)
            Trigger_Index = Alert_header.index('Trigger Alert Number')    
            trigger_value = Alert_Sheet.cell(2, Trigger_Index+1).value  
            Trigger_Num = int(alert_results[0][-1])
            if int(Trigger_Num)> 0:
                Alert_Sheet.batch_clear(['A2:C'])
                Alert_Sheet.update(values= alert_results, range_name='A2:C')


                if (Trigger_Num) >= int(trigger_value):
                    if not send_slack_notification(alert_results):        
                        ui.Status_label.setText("Can't send Slack Alert")
                        QApplication.processEvents()

            Last_Alert__Run = time.time()
            return True
        except Exception as e:
            messagebox.showerror("Error",e)
            # return False
        # return True


# MCDR_ALert_Function = threading.Thread(target=MCDR_ALert)
# MCDR_ALert_Function.start()    



def Query_Google_Sync():

    global pause,Timer,LastUpdateDate,SymbolCode,QTY,CustodianCode,UnifiedCode,comment,MCDR_Code,cursor,Alert_Sheet,Last_Alert__Run
    MCDR_Code = 0

    while  pause ==1:
        if pause == 0 :
            break
        elif pause == 2 :
            return 
    # ui.User_N_input.setText ('farid.shawky@thndr.app')
    # password = ui.PW_input.setText ('8mrM7dZF@5LP')
    # ui.Seconds_num.setText ('22')

    server = ui.Server_input.text()
    database = ui.Db_input.text()
    username = ui.User_N_input.text()
    password = ui.PW_input.text()
    Timer =  ui.Seconds_num.text()
    Start_line = ui.From_row.text()
    End_Line = ui.To_row.text()

    if ui.Seconds_num.text() == '' :
       error = 4
       return error

    if stop_flag.is_set():
        threading.Event.clear(stop_flag)
        image_detection_thread = threading.Thread(target=detect_and_click)
        image_detection_thread.start()

        
#------------------------------------------------------------------------------------ google sheet ----------------------------------------------------

    # Set up authentication credentials
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_file = 'creds.json'
    google_jsonfile = 'googlekey.json'
    try:
        creds = None
        if os.path.exists(creds_file):           # we check if the file path is exist or not
                with open('creds.json', 'r') as f:    # we open the creds file as in read mode and in (f) as a handle of file and load it to creds_json
                    creds_json = f.read()
                creds = ServiceAccountCredentials.from_json(creds_json)     # load the creds that will use to access to google sheet from creds_json
                ui.Status_label.setText("Try to Open Google Sheet...")
                QApplication.processEvents()
        else:
                if not os.path.exists(google_jsonfile):
                    ui.Status_label.setText("The Json file does not exist.")
                    QApplication.processEvents()
                    error = 3
                    return error
                else:
                    creds = ServiceAccountCredentials.from_json_keyfile_name(google_jsonfile, scope)   # if creds.json file not exist we get the creds by googlekey.json file as usual
                    with open(creds_file , 'w') as f:
                        f.write(creds.to_json())


        client = gspread.authorize(creds)


        if os.path.exists('Sheet_Url.txt'): 
                with open('Sheet_Url.txt', 'r') as u:   #--------------read URL for google sheet
                        sheet_url = u.read()

                #Access the Google Sheet
                sheet = client.open_by_url(sheet_url).worksheet('Orders')
                Alert_Sheet = client.open_by_url(sheet_url).worksheet('MCDR Alert')
        else:
                pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
                return

        # Find the last row with values
        last_row = len(sheet.col_values(1)) + 1

        # get the index of the header
        header =sheet.row_values(1)
        LastUpdateDate_index = header.index('LastUpdateDate')
        # Alert_header = Alert_Sheet.row_values (1)
        # Trigger_Index = Alert_header.index('Trigger Alert Number')

        if sheet.cell(last_row-1, LastUpdateDate_index+1).value !='LastUpdateDate' :
                LastUpdateDate = sheet.cell(last_row-1, LastUpdateDate_index+1).value  
    except Exception as d:
       messagebox.showinfo("Error",d)   
       return False
#-------------------------------------------------------------------------------------AZURE----------------------------------- 
    if pause == 2 :
        pause = 0
        return

    try:
        with open('Query.txt', 'r') as f:
         SQL_Query = f.read()
         SQL_Query = SQL_Query.replace('--*',LastUpdateDate )    #----------------------------------------- 
        # with open ('Alert_Query.txt','r') as m:
        #     Alert_Query  = m.read()
        #     Alert_Query = Alert_Query.replace('--*',f'-{Timer}' )
               
    except Exception as e:
        ui.Status_label.setText(e)
        QApplication.processEvents()        
        return False

    # Establish connection to Azure SQL Database
    try:
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        conn = pyodbc.connect(conn_str)
    except Exception as d:
        messagebox.showinfo("Error",d)
        return False

    # Execute the SQL query and fetch the results
    conn.commit()
    cursor = conn.cursor()
    cursor.execute(SQL_Query)
    query_results = cursor.fetchall()

   

    # Write the query results to the CSV file with proper encoding
    filename = 'Q_results.csv'
    with open(filename, 'w', newline='', encoding='utf-16') as csvfile:
        writer = csv.writer(csvfile)
        # Get the column names from the cursor's description
        column_names = [column[0] for column in cursor.description]
        column_names.append("Eligibility")
        column_names.append("Action")
        


        # Write each row of the query results
        writer.writerows(query_results)

     
    # cursor.execute(Alert_Query)
    # alert_results = cursor.fetchall()

    # Close the database connection
    # conn.close()

#-------------------------------------------------------------------------------------AZURE------------------------- 

    SQL_Query.replace("'", "")
   
    try:    
        if  os.path.getsize(filename) > 0:
            
            df = pd.read_csv(filename, encoding='utf-16', header=None)   
            if header != column_names:
                sheet.insert_rows([column_names], row=1)

            
            else:
                # sheet.update(f'A{last_row}', df.map(str).values.tolist()) 
                sheet.update(values=df.map(str).values.tolist(), range_name=f'A{last_row}')  
    
    except Exception as d:
        messagebox.showerror("Error",d)
        # time.sleep(5)
        return False   
        # pass


    if not MCDR_ALert():
        return False

    header = sheet.row_values(1)
    T_Type_index = header.index('T-Type')
    CustodianCode_index = header.index('CustodianCode')
    OrgQty_index = header.index('OrgQty')
    SymbolCode_index = header.index('SymbolCode')
    UnifiedCode_index = header.index('UnifiedCode')
    Action_index = header.index('Action')
    Eligibility_index =  header.index('Eligibility')

    
    ui.Status_label.setText("Google Sheet updated")
    QApplication.processEvents()

    #filename.close()

    Sheet2 = client.open_by_url(sheet_url).worksheet('MCDR Mapping')

    # find the unresolved  issues rows
    # google_rows = sheet.get_all_values()
    google_rows = sheet.get_values(f'A1:Z{last_row}')

    if isinstance(Start_line, str) and not Start_line.isdigit() or Start_line == "" :  
        Start_line = 2

    if isinstance(End_Line, str) and not End_Line.isdigit() or End_Line == "" :
        End_Line = last_row

  
    
    for i, row in enumerate(google_rows):               # the main loop that excute the boobking steps-----------------------

        comment = ''
        MCDR_Code = 0
        while  pause ==1:
                if pause == 0 :
                    break
                elif pause == 2 :
                  return

        if row[Action_index] == '' and row[Eligibility_index] == '1' and i>= int(Start_line)-1 and i<= int(End_Line)-1   :       #   and i > 1430   and i > ???           

            SymbolCode = row[SymbolCode_index]
            UNI_Code = row[UnifiedCode_index]
            QTY =  row[OrgQty_index]
            CustodianCode = row[CustodianCode_index]
            UnifiedCode = row[UnifiedCode_index]
            Stock_Num = 0
           
            if row[T_Type_index] == 'T0Sell' and  ui.T0.isChecked():   #---------- T0 function --------------
                try:
                   cell = Sheet2.find(SymbolCode)
                   row_num = cell.row
                   row_sheet2 = Sheet2.row_values(row_num)
                   MCDR_Code = row_sheet2[3]
                   if not  T0_call(MCDR_Code,UNI_Code,sheet,i,Action_index):
                     return False  
                except Exception as q:
                    comment = 'Can not find MCDR Number' 
                    MCDR_Code = 0       
            
                sheet.update_cell(i+1, Action_index+1, comment)
                
            if row[T_Type_index] == 'T1Sell' and   ui.T1.isChecked():            #---------- T1 function --------------
                try:
                   cell = Sheet2.find(SymbolCode)
                   row_num = cell.row
                   row_sheet2 = Sheet2.row_values(row_num)
                   MCDR_Code = row_sheet2[3]
                   if not T1_call(MCDR_Code,UNI_Code,sheet,i,Action_index):
                      return False 
                except Exception as q:
                    comment = 'Can not find MCDR Number' 
                    MCDR_Code = 0    
                
                sheet.update_cell(i+1, Action_index+1, comment)                   

            if row[T_Type_index]=='Sell' and  ui.T2_unbooking.isChecked():
                   try:
                        cell = Sheet2.find(SymbolCode)
                        row_num = cell.row
                        row_sheet2 = Sheet2.row_values(row_num)
                        MCDR_Code = row_sheet2[3]
                        result = Un_Booking() 
                        if result is False:
                            return False
                        else:
                            pass
                   except Exception as q:
                        comment = 'Can not find MCDR Number' 
                        MCDR_Code = 0 
                                                   
            if row[T_Type_index] == 'Sell'  and   ui.T2.isChecked() :
                try:
                   cell = Sheet2.find(SymbolCode)
                   row_num = cell.row
                   row_sheet2 = Sheet2.row_values(row_num)
                   MCDR_Code = row_sheet2[3]
                   if not Booking_T2():
                      return False
                except Exception as q:
                    comment = 'Can not find MCDR Number' 
                    MCDR_Code = 0 
                sheet.update_cell(i+1, Action_index+1, comment)    
                
            if int(time.time()-Last_Alert__Run)>=30:
                ui.Status_label.setText(f'Running Alert Query..')
                QApplication.processEvents()
                MCDR_ALert()
               
        
    return True


def pause_execution():
    global pause
    global image_detection_thread
    if pause == 0:
        pause = 1 
        threading.Event.set(stop_flag)
        image_detection_thread.join()

        answer = pyautogui.confirm("Do you want to Exit ?", "PAUSE", buttons=["Yes", "No"])
        if answer == "Yes":

            pause = 2
            #threading.Event.set(stop_flag)
            ui.Status_label.setText("Operations have been cancelled..")
            QApplication.processEvents()
            return
        elif answer == "No":
            threading.Event.clear(stop_flag)
            image_detection_thread = threading.Thread(target=detect_and_click)
            image_detection_thread.start()

            pause = 0
           
            return


# Set the hotkey
keyboard.add_hotkey('ALT+ p', pause_execution)




def Run_Function():
    sleep_duration = 0
    global Timer,Last_Alert__Run,pause

    #keyboard.add_hotkey('esc', pause_execution)
    Timer = ui.Seconds_num.text()

    while True:
        Last_Run = int(time.time())
        ui.Status_label.setText("Start running the query")
        QApplication.processEvents()
        if pause == 2 :
            pause = 0
            return
        # error = Query_Google_Sync()
        if not Query_Google_Sync():
            return False

        else:
            Current_time = int(time.time())
            dev = int ( Current_time - Last_Run )
            WinActivate(app_title,1)
            sleep_duration = int(Timer) - dev

            if sleep_duration > 0:
                for i in range(sleep_duration):
                    while pause == 1:
                        QApplication.processEvents() 
                    ui.Status_label.setText(f'wait for next run {sleep_duration-i} Sec')
                    QApplication.processEvents()
                    time.sleep(1)
                    if int(time.time()-Last_Alert__Run)>=30:
                        ui.Status_label.setText(f'Running Alert Query..')
                        QApplication.processEvents()
                        MCDR_ALert()

                    if pause == 2 :
                        pause = 0
                        ui.Status_label.setText(f'Operations have been cancelled..')
                        QApplication.processEvents()
                        return





def check_pause():
    global pause 
    if pause == 2:
        pause = 0
        pause_execution()
    elif pause == 0:
        pause_execution()
        

    

if __name__ == "__main__":
    #import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Win.Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.setWindowTitle(app_title)
    ui.Server_input.setText("10.70.4.100")     #put value in server IP
    ui.Db_input.setText("master")     #put value in Database Name
    ui.PW_input.setEchoMode(QtWidgets.QLineEdit.Password)
    ui.Status_label.setAlignment(QtCore.Qt.AlignCenter)
    ui.T0.setChecked(True)
    ui.T1.setChecked(True)

    ui.start_b.clicked.connect(Run_Function)
    ui.puase_b.clicked.connect(check_pause)

    # Set up a keyboard shortcut to call the pause_execution function
    #shortcut = QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.ALT + QtCore.Qt.Key_P), MainWindow)
    #shortcut.activated.connect(pause_execution)


    #keyboard.add_hotkey('esc', pause_execution)

    MainWindow.show()
    sys.exit(app.exec_())

