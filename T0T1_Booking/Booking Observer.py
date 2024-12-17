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
import mainwindow3 as Win                                      # import the file of the window and all interface
from pywinauto import Application
import pyperclip
import keyboard
import threading
import pandas as pd
#import pytesseract

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


Last_Run = 0
Timer = 0
pause = 0
LastUpdateDate = '2023-12-31 10:00:00.973000'
issue_flag = 0
app_title = 'Booking Observer'
SymbolCode = ''
QTY = 0
MCDR_Code = 0
CustodianCode = ''
UnifiedCode = ''
comment = ''
Stock_Num = 0





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
        # Check if the image appears on the screen
        if pyautogui.locateOnScreen(path+'Issue_Num.png') is not None :
            issue_flag = 1
            time.sleep(0.13)
            # Perform the click action on the image
            pyautogui.click(path+'Issue_Num_Ok.png')
            #break
             

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



def T0_call(MCDR_Code,UNI_Code):

    global pause
    global T0
    global issue_flag
    issue_flag = 0
    global SymbolCode

    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return

    if not WinActivate(T0,3):
       #pyautogui.alert("Can't Find T+0 ","Error")
       Stock_Num = 'not_found'
       return Stock_Num

    
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
                Stock_Num = 'MAx QTY'
                return Stock_Num
            elif img == 2:
                time.sleep(0.3)
                clickOn("done.png", 0, 1,10)
                return Stock_Num 

        else: 
              clickOn('back.png',0,1,15)
              time.sleep(0.15)    
              Stock_Num = 'MAx_QTY'
              return Stock_Num
    if img ==2:
            clickOn("done.png",0,1,30)
            Stock_Num = 'Sold or No Available Qty' 
            return Stock_Num 
 



def T1_call(MCDR_Code,UNI_Code):
    global pause
    global issue_flag
    global SymbolCode
    issue_flag = 0

    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return


    if not WinActivate(T1,3):
       #pyautogui.alert("Can't Find T+0 ","Error")
       Stock_Num = 'not_found'
       return Stock_Num

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
        clickOn('No_Issue_Select.png',0,1,70)
        issue_flag = 0
        time.sleep(.4)
    if issue_flag == 1:
        time.sleep(.3)
        clickOn('No_Issue_Select.png',0,1,70)
        time.sleep(.4)
    
        issue_flag = 0

   
  
    clickOn("search.png",0,1,20)

    time.sleep(0.2)
    clickOn('Boked.png', 0, 1,20)
    time.sleep(0.2)
    x,y,img = Detect_img('record_sell_order.png',0,0,15)
    if img == None:
      Stock_Num = None
      return Stock_Num 

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

       

    else:
       Stock_Num = "None"
       clickOn('back.png',0,1,20) 
       time.sleep(0.2)

    return Stock_Num




def Un_Booking():
    global pause
    global UnifiedCode
    global Symbol_Num
    global Stock_Num
    global CustodianCode

    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return

    if not WinActivate(Book_Cancel,3):
       error = 1
       Stock_Num = 0 
       return error,Stock_Num
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
                    error = 0
                    return error,Stock_Num
                elif img == 2:
                    time.sleep(0.2)
                    clickOn('FRM_40401_Ok.png',0,1,8)
                    time.sleep(0.4)
                    Stock_Num = 0
                    error = 3
                    return error,Stock_Num
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
                        error = 3
                        return error,Stock_Num

                    elif img == 2:
                        #clickOn('Delete_Done.png',0,1,40)
                        time.sleep(0.3)
                        clickOn('Delete_Done_ok.png',0,1,40)
                        time.sleep(0.2)
                        error = 0
                        return error,Stock_Num
                    elif img == 3 :
                        time.sleep(0.3)
                        clickOn('FRM_40401_Ok.png',0,1,10) 
                        time.sleep(0.3)
                        x,y,img = Detect_img('Order_Is_Deleted.png','Delete_Done.png',0,30)
                        if  img == 1:
                            time.sleep(0.35)
                            clickOn('Order_Is_Deleted_ok.png',0,1,10)
                            time.sleep(0.2)
                            error = 0
                            return error,Stock_Num
                        elif img == 2:
                            time.sleep(0.35)
                            clickOn('Delete_Done_Ok.png',0,1,10)
                            time.sleep(0.2)
                            error = 0
                            return error,Stock_Num
                        else:

                            error = 3
                            Stock_Num = 0
                            return error,Stock_Num
                            
                if img == None:
                        Stock_Num  = 0
                        error = 3
                        return error,Stock_Num
            else:
                error = 2
                Stock_Num = 0
                return error,Stock_Num



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


    while True:
        if pause == 0 :
            break
        elif pause == 2 :
            return
  
    if not WinActivate(T2,4):
       error = 11
       return error
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
            error = 12
            return  error
        elif img == 2 :
            time.sleep(0.2)
            clickOn('Qty_Is_Booked_Ok.png',0,1,40)
            error = 0
            return error
        elif img == 3 :
            time.sleep(0.2)
            clickOn('Cannot_delete_ok.png',0,1,10)
            time.sleep(0.2)
            error = 13
            return error




def Query_Google_Sync():

    global pause
    global Last_Run
    global Timer
    global LastUpdateDate
    global SymbolCode
    global QTY
    global CustodianCode
    global UnifiedCode
    global comment 
    global MCDR_Code
    MCDR_Code = 0


    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return 

    server = ui.Server_input.text()
    database = ui.Db_input.text()
    username = ui.User_N_input.text()
    password = ui.PW_input.text()
    Timer =  ui.Seconds_num.text()
    if ui.Seconds_num.text() == '' :
       error = 4
       return error

    Last_Run = time.time()

    if stop_flag.is_set():
        threading.Event.clear(stop_flag)
        image_detection_thread = threading.Thread(target=detect_and_click)
        image_detection_thread.start()

        
#------------------------------------------------------ google sheet ----------------------------------------------------

    # Set up authentication credentials
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_file = 'creds.json'
    google_jsonfile = 'googlekey.json'

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
    else:
        pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
        return

    # Find the last row with values
    last_row = len(sheet.col_values(1)) + 1

    # get the index of the header
    header = sheet.row_values(1)
    LastUpdateDate_index = header.index('LastUpdateDate')

    if sheet.cell(last_row-1, LastUpdateDate_index+1).value !='LastUpdateDate' :
        LastUpdateDate = sheet.cell(last_row-1, LastUpdateDate_index+1).value    
        
#-------------------------------------------------------------------------------------------AZURE-------------- 
    # Azure SQL Database connection details
    
    try:
        with open('Query.txt', 'r') as f:
         SQL_Query = f.read()
         
         SQL_Query = SQL_Query.replace('--*',LastUpdateDate )    #-----------------------------------------

        
    except Exception as e:
        error = 1
        return error

    # Establish connection to Azure SQL Database
    try:
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        conn = pyodbc.connect(conn_str)
    except Exception as e:
        error = 2
        return error

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


    csvfile.close()   
    


    # Close the database connection
    conn.close()


#-------------------------------------------------------------------------------------------AZURE-------------- 


    # Read the CSV file
    #filename = 'Q_results.csv'
    #with open(filename, 'r', encoding='utf-16') as csvfile:
     #   reader = csv.reader(csvfile)
      #  rows = list(reader)
        


    SQL_Query.replace("'", "")
   


    try:    
        df = pd.read_csv(filename, encoding='utf-16', header=None)   
        
        if header != column_names:
          sheet.insert_rows([column_names], row=1)
        else:
            sheet.update(f'A{last_row}', df.map(str).values.tolist())  
    except Exception as q:
        pass

    
    # refresh the index of the header after write it
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
    google_rows = sheet.get_all_values()
    #google_rows = filter(lambda row: row[Action_index] == '' and row[Eligibility_index] == '1' and (row[T_Type_index] == 'T0Sell' or row[T_Type_index] == 'T1Sell'), sheet.get_all_values())


    for i, row in enumerate(google_rows):               # the main loop that excute the boobking steps-----------------------

        comment = ''
        MCDR_Code = 0

        while  True:
                if pause == 0 :
                    break
                elif pause == 2 :
                  return


        if row[Action_index] == '' and row[Eligibility_index] == '1' :             # Skip header row

            
            SymbolCode = row[SymbolCode_index]
            UNI_Code = row[UnifiedCode_index]
            QTY =  row[OrgQty_index]
            CustodianCode = row[CustodianCode_index]
            UnifiedCode = row[UnifiedCode_index]

            

            
            if row[T_Type_index] == 'T0Sell'and ui.T0.isChecked():

            
               try:
                   cell = Sheet2.find(SymbolCode)
                   row_num = cell.row
                   row_sheet2 = Sheet2.row_values(row_num)
                   MCDR_Code = row_sheet2[3]
               except Exception as q:
                   sheet.update_cell(i+1, Action_index + 1 , 'Can not find MCDR Number')
                   MCDR_Code = 0

               if MCDR_Code !=0:
                   
                   Stock_Num =  T0_call(MCDR_Code,UNI_Code)
                   if Stock_Num is not None and Stock_Num.isnumeric(): 
                      number = 0  
                      number = int(Stock_Num)
                      if number > 0  :
                         
                          sheet.update_cell(i+1, Action_index+1, Stock_Num + '  Booked')
                   elif Stock_Num =='not_found' :
                      error = 3
                      return error
                   else:
                      sheet.update_cell(i+1, Action_index+1, Stock_Num)


            if row[T_Type_index] == 'T1Sell' and ui.T1.isChecked():


               try:
                   cell = Sheet2.find(SymbolCode)
                   row_num = cell.row
                   row_sheet2 = Sheet2.row_values(row_num)
                   MCDR_Code = row_sheet2[3]
               except Exception as q:
                   sheet.update_cell(i+1, Action_index + 1 , 'Can not find MCDR Number')
                   MCDR_Code = 0

               if MCDR_Code !=0:
                   
                   Stock_Num =  T1_call(MCDR_Code,UNI_Code)
                   
                   if Stock_Num is not None and Stock_Num.isnumeric():
                      number = 0  
                      number = int(Stock_Num)
                      if number > 0  :

                        sheet.update_cell(i+1, Action_index+1, Stock_Num + ' ' + 'Booked')
                   elif Stock_Num =='not_found' :
                      error = 3
                      return error

                   else:
                      sheet.update_cell(i+1, Action_index+1, 'No Stock')

            if row[T_Type_index]=='Sell' and ui.T2_unbooking.isChecked():

               try:
                   cell = Sheet2.find(SymbolCode)
                   row_num = cell.row
                   row_sheet2 = Sheet2.row_values(row_num)
                   MCDR_Code = row_sheet2[3]
               except Exception as q:
                   sheet.update_cell(i+1, Action_index + 1 , 'Can not find MCDR Number')
                   MCDR_Code = 0
                   

               if MCDR_Code !=0:

                   error,Stock_Num = Un_Booking()

                   if error == 1 :
                        ui.Status_label.setText("Can't open MCDR System ..")
                        QApplication.processEvents()
                        error = 3
                        return error
                   elif error == 0:
                        comment = Stock_Num + ' canceled'
                       #Sheet2.update_cell(i+1, Action2_index+1, Stock_Num) 
                   elif error ==3 :
                        comment = 'Qty in MkT '

                   else:
                        comment = comment + '0 stock canceled '
                   if pause == 2 :
                        return    
               

            if row[T_Type_index] == 'Sell' and ui.T2.isChecked() :

                error = Booking_T2()
                if error == 11:
                   return error 
                elif error == 0 :
                    comment = comment  + QTY + ' Booked'
                   #Sheet2.update_cell(i+1, Action2_index+1,'Booked '+ QTY)
                elif error == 12 :
                    comment = comment + '- No Balance for ' + QTY
                elif error == 13:
                    comment = comment + ' Code Suspended  '    

                sheet.update_cell(i+1, Action_index + 1, comment)
                comment = '' 
                if pause == 2 :
                  return             




def pause_execution():
    global pause
    global image_detection_thread
    if pause == 0:
        pause = 1 
        threading.Event.set(stop_flag)
        image_detection_thread.join()

        answer = pyautogui.confirm("Do you want to Exit ?", "PAUSE", buttons=["Yes", "No"])
        if answer == "Yes":
            print("yes pressed")
            
            pause = 2
            #threading.Event.set(stop_flag)
            ui.Status_label.setText("Operations have been cancelled..")
            QApplication.processEvents()
            return
        elif answer == "No":
            print("no pressed")
            threading.Event.clear(stop_flag)
            image_detection_thread = threading.Thread(target=detect_and_click)
            image_detection_thread.start()

            pause = 0
           
            return


# Set the hotkey
keyboard.add_hotkey('ALT+ p', pause_execution)




def Run_Function():
    sleep_duration = 0
    global Timer
    global Last_Run
    global pause
    #keyboard.add_hotkey('esc', pause_execution)
    Timer = ui.Seconds_num.text()

    while True:
  
        ui.Status_label.setText("Start running the query")
        QApplication.processEvents()
        error = Query_Google_Sync()
        if error == 1:
           ui.Status_label.setText("Can't Find Query File...")
           QApplication.processEvents()
           return
        elif error == 2:
           ui.Status_label.setText("Error in connecting SQL Server")
           QApplication.processEvents()
           return
        elif error == 3:
           ui.Status_label.setText("Can't Find MCDR")
           QApplication.processEvents()
           break
        elif error == 4:
           ui.Status_label.setText("Please enter the number of seconds..")
           QApplication.processEvents()
           return
        elif error == 5 :
            ui.Status_label.setText("Cannot find some Data File..")
            QApplication.processEvents()
            return
        elif pause == 2:
            pause = 0
            break

        
        else:
            Current_time = time.time()
            dev = float ( Current_time - Last_Run )
            #WinActivate(app_title,1)
            print(dev)
            ui.Status_label.setText("wait for next run....")
            QApplication.processEvents()
            sleep_duration = float(Timer) - dev
            print(sleep_duration)
            if sleep_duration > 0:
                #Timer = Main_Timer
                time.sleep(sleep_duration)
            

        while  True:
            if pause == 0 :
                break
            elif pause == 2 :
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


