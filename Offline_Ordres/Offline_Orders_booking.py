import pyodbc
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import csv
import codecs  # Import the codecs module for encoding support
#from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtWidgets, QtGui, QtCore

import datetime
from decimal import Decimal
import pyautogui
import time
import Booking_window2 as Win                                      # import the file of the window and all interface
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

Last_Run = 0
Timer = 0
pause = 0
LastUpdateDate = '2023-12-31 10:00:00.973000'
issue_flag = 0
app_title = 'Offline Ordres'
SymbolCode = ''
UnifiedCode = ''
comment = ''

Order_Type_Index=0
Order_Type= 0
Market_Code_index=0
Market_Code=0
ISIN_Code_index=0
ISIN_Code=0
MCDR_Code = 0
QTY_index=0
QTY=0
Price_Index=0
Price=0
Thndr_Code_Index=0
Thndr_Code=0
CustodianCode_index=0
CustodianCode=0
Purchase_Power_Indix=0
Purchase_Power = 0
T2_Booked_Index = 0
T2_Booked=0
T1_Booked_Index = 0
T1_Booked=0
T0_Booked_Index = 0
T0_Booked = 0
Offline_Booking_indx = 0
Offline_Booking = 0

X_stream_Index = 0
X_stream = 0
Stock_Num = ''

All_Even_Odd = 1
pause =0
Last_Run=0  
Action1_index =0
comment=''
filename = 'Q_results.csv' 
autoit_script = "execute.au3"




#pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Thndr\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
pyautogui.PAUSE = 0.1

path = os.getcwd() + '\\data\\'


def Change_language(hex_code):
    if SendMessage( GetForegroundWindow(), WM_INPUTLANGCHANGEREQUEST, 0, hex_code) == 0:
       SendMessage( GetForegroundWindow(), WM_INPUTLANGCHANGEREQUEST, 0, hex_code) 



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



def T0_call():
    global pause
    global T0
    global UnifiedCode
    global MCDR_Code
    global Stock_Num
    global comment
    global issue_flag
    issue_flag = 0

    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return

    comment = ''
    Stock_Num = ''
    if not WinActivate(T0,3):
       #pyautogui.alert("Can't Find T+0 ","Error")
       comment = 'MCDR Error'
       return Stock_Num

    
    Change_language(0x4090409) 
    time.sleep(0.1)
    clickOn("client.png",90,2,5)
    time.sleep(0.1)
    pyautogui.typewrite(UnifiedCode)
    clickOn("stock.png",90,2,5)
    pyautogui.typewrite(MCDR_Code)
    time.sleep(0.2)
    clickOn("search.png",0,1,5)
    time.sleep(0.1)
    if issue_flag == 1:
        time.sleep(1.3)
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
                return Stock_Num
            elif img == 2:
                time.sleep(0.3)
                clickOn("done.png", 0, 1,10)
                comment =''
                return Stock_Num 

        else: 
              clickOn('back.png',0,1,15)
              time.sleep(0.15)    
              comment = 'MAx_QTY'
              return Stock_Num
    if img ==2:
            clickOn("done.png",0,1,30)
            comment = 'Sold or No Available Qty' 
            return Stock_Num 
 

def T1_call():
    global pause
    global issue_flag
    global UnifiedCode
    global ISIN_Code
    global MCDR_Code
    global Stock_Num
    global comment
    issue_flag = 0

    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return


    comment = '' 
    Stock_Num = ''       
    if not WinActivate(T1,3):
       #pyautogui.alert("Can't Find T+0 ","Error")
       comment = 'MCDR Error'
       return Stock_Num

    Change_language(0x4090409) 
    time.sleep(0.1)
    clickOn("client.png",90,2,20)
    time.sleep(0.1)
    pyautogui.typewrite(UnifiedCode)
    clickOn("SymbolCode.png",90,2,20)
    pyautogui.typewrite(ISIN_Code)
    time.sleep(0.2)
    clickOn("search.png",0,1,20)
    
    time.sleep(0.4)

    if issue_flag == 1:

        clickOn('No_Issue_Select.png',0,1,700)
        time.sleep(0.15)
        issue_flag = 0

   
  
    clickOn("search.png",0,1,20)

    time.sleep(0.2)
    clickOn('Boked.png', 0, 1,20)
    time.sleep(0.2)
    x,y,img = Detect_img('record_sell_order.png',0,0,15)
    if img == None:
      comment = 'No Qty'  
      Stock_Num = ''
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
       comment = ''

    else:
       comment = '0 Qty' 
       Stock_Num = ''
       clickOn('back.png',0,1,20) 
       time.sleep(0.2)

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
    global ISIN_Code
    global Stock_Num


    while True:
        if pause == 0 :
            break
        elif pause == 2 :
            return
    
    comment = ''
    Stock_Num = ''

    if not WinActivate(T2,4):
       comment  = 'MCDR Error'
       return Stock_Num

    else:

        time.sleep(0.2)
        clickOn('SymbolCode.png',120,2,30)
        #Change_language() 
        keyboard2.type(ISIN_Code)

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
        time.sleep(0.4)
        pyautogui.typewrite('\t') 
        time.sleep(0.4)
       
        keyboard2.type(Expirydate)
        pyautogui.typewrite('\t') 
        clickOn('Save.png',0,1,30)

        x,y,img = Detect_img ('Not_Enough_Balance.png','Qty_Is_Booked.png','not_listed.png',40)
        if img == 1: 
            clickOn('Delete_Done_Ok.png',0,1,20)
            comment  = 'Not Enough Balance'
            Stock_Num = ''
            return  Stock_Num
        elif img == 2 :
            time.sleep(0.2)
            clickOn('Qty_Is_Booked_Ok.png',0,1,40)
            comment = ''
            Stock_Num = QTY
            return Stock_Num
        elif img == 3 :
            time.sleep(0.2)
            clickOn('Cannot_delete_ok.png',0,1,10)
            time.sleep(0.2)
            comment = 'Code Suspended'
            return Stock_Num



def Query_Google_Sync():

    global Order_Type_Index
    global Order_Type
    global Market_Code_index
    global Market_Code
    global ISIN_Code_index
    global ISIN_Code
    global MCDR_Code
    global QTY_index
    global QTY
    global Price_Index
    global Price
    global Thndr_Code_Index
    global Thndr_Code
    global CustodianCode_index
    global CustodianCode
    global Purchase_Power_Indix
    global Purchase_Power
    global T2_Booked_Index 
    global T2_Booked
    global T1_Booked_Index 
    global T1_Booked
    global T0_Booked_Index 
    global T0_Booked
    global Offline_Booking_indx
    global Offline_Booking
    
    global X_stream_Index
    global X_stream 
        
    global pause
    global Last_Run
    global Action1_index 
    global Ticker_index 
    global comment
    global autoit_script
    global filename 
    global UnifiedCode

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

    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return

    Last_Run = time.time()

    ui.Status_label.setText("Try to Open Google Sheet...")
    QApplication.processEvents()

    # Set up authentication credentials
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_file = 'creds.json'
    google_jsonfile = 'offline-orders-booking.json'

    creds = None
    if os.path.exists(creds_file):           # we check if the file path is exist or not
       with open('creds.json', 'r') as f:    # we open the creds file as in read mode and in (f) as a handle of file and load it to creds_json
        creds_json = f.read()
       creds = ServiceAccountCredentials.from_json(creds_json)     # load the creds that will use to access to google sheet from creds_json
       
    else:
       if not os.path.exists(google_jsonfile):
          error = 2
          return error
       else:
          creds = ServiceAccountCredentials.from_json_keyfile_name(google_jsonfile, scope)   # if creds.json file not exist we get the creds by googlekey.json file as usual
          with open(creds_file , 'w') as f:
            f.write(creds.to_json())

    ui.Status_label.setText("Try to Open Google Sheet...")
    QApplication.processEvents()  

    client = gspread.authorize(creds)

    if os.path.exists('Sheet_Url.txt'): 
        with open('Sheet_Url.txt', 'r') as u:   #--------------read URL for google sheet
            sheet_url = u.read()

        #Access the Google Sheet
        sheet = client.open_by_url(sheet_url).worksheet('Orders')
    else:
        pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
        QApplication.processEvents()
        error = 3
        return error 


    # get the index of the sheet1_header

    sheet1_header = sheet.row_values(1)

    Order_Type_Index = sheet1_header.index('Type')
    Market_Code_index = sheet1_header.index('Market Code')
    ISIN_Code_index = sheet1_header.index('ISIN')
    QTY_index = sheet1_header.index('Qty')
    Price_Index = sheet1_header.index('Price')
    UnifiedCode_index = sheet1_header.index('UnifiedCode')

    Thndr_Code_Index = sheet1_header.index('Thndr Code')
    CustodianCode_index = sheet1_header.index('Custody')
    Purchase_Power_Indix = sheet1_header.index('Purchase Power')
    T2_Booked_Index = sheet1_header.index('T2 Booked')
    T1_Booked_Index = sheet1_header.index('T1 Booked')
    T0_Booked_Index = sheet1_header.index('T0 Booked')
    Offline_Booking_indx = sheet1_header.index('Offline Booking')
    X_stream_Index = sheet1_header.index('X_stream')
    Comment_indx = sheet1_header.index('Comment')


    
    

    ui.Status_label.setText("Sync Google Sheet..")
    QApplication.processEvents()
    #filename.close()

    

    sheet1_row = sheet.get_all_values()

    #with open("Input_var.csv", 'w', newline='', encoding='utf-16') as csvfile:                
    #     writer = csv.writer(csvfile)
    #     #writer.writerows(sheet1_row)

    Sheet2 = client.open_by_url(sheet_url).worksheet('Code')

      
    for i, row in enumerate(sheet1_row):               # the main loop that excute the boobking steps-----------------------

        MCDR_Code = 0


        while  True:
                if pause == 0 :
                    break
                elif pause == 2 :
                  return

        if row[Offline_Booking_indx] == '' and row[Order_Type_Index] == 'Sell':         
            
            ISIN_Code = row[ISIN_Code_index]
            QTY =  row[QTY_index]
            CustodianCode = row[CustodianCode_index]
            UnifiedCode = row[UnifiedCode_index]  


           


            #print(row[CustodianCode_index] + ' ' + row[Offline_Booking_indx])      
            if ( row[CustodianCode_index] == '4625' or row[CustodianCode_index] == '4503')  and row[X_stream_Index] == 'Qty Not Booked'  and ui.T2.isChecked() and row[Offline_Booking_indx] == '' and row[Comment_indx] =='' :            # Skip header row

                
                Stock_Num = Booking_T2()

                sheet.update_cell(i+1, Offline_Booking_indx + 1, Stock_Num)
                sheet.update_cell(i+1, Comment_indx + 1, comment)
                
                comment = ''     
               
                
            if ( row[CustodianCode_index] == '46257' or row[CustodianCode_index] == '45037')  and ui.T1.isChecked()  and row[Offline_Booking_indx] == '' and row[Comment_indx] =='' :             # Skip header row

                    
                   Stock_Num =  T1_call ()

                   sheet.update_cell(i+1, Offline_Booking_indx + 1, Stock_Num)
                   sheet.update_cell(i+1, Comment_indx + 1, comment)
                   
                 
            if ( row[CustodianCode_index] == '46256' or row[CustodianCode_index] == '45036')  and ui.T0.isChecked() and row[Offline_Booking_indx] =='' and row[Comment_indx] == '' :      
                  

                try:
                   cell = Sheet2.find(ISIN_Code)
                   row_num = cell.row
                   row_sheet2 = Sheet2.row_values(row_num)
                   MCDR_Code = row_sheet2[3]
                except Exception as e:
                   sheet.update_cell(i+1, Comment_indx +1, 'Can not  find MCDR Number')

                if MCDR_Code != 0 :   
                   Stock_Num = T0_call()

                   sheet.update_cell(i+1, Offline_Booking_indx + 1, Stock_Num)
                   sheet.update_cell(i+1, Comment_indx + 1, comment)
                   
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
            
            pause = 2
            #threading.Event.set(stop_flag)
            ui.Status_label.setText("Operations have been cancelled..")
            QApplication.processEvents()
            return
        elif answer == "No":
            threading.Event.clear(stop_flag)
            image_detection_thread = threading.Thread(target=detect_and_click)
            image_detection_thread.start()


            #detect_and_click()
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
            WinActivate(app_title,1)
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


