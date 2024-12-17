import pyodbc
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import csv
import codecs  # Import the codecs module for encoding support
#from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtWidgets, QtGui, QtCore
from pynput.keyboard import Controller, Key 
import datetime
from decimal import Decimal
import pyautogui
import time
import Booking_window as Win                                      # import the file of the window and all interface
from pywinauto import Application
import pyperclip
import keyboard
import threading
import ctypes
#import pytesseract
import win32api
import os
import json
import sys
#import tkinter as tk
#from tkinter import messagebox
#from dateutil.relativedelta import relativedelta
import win32clipboard

from win32con import WM_INPUTLANGCHANGEREQUEST
from win32gui import GetForegroundWindow
from win32api import SendMessage


current_date = current_date = datetime.datetime.now()


new_date = current_date + datetime.timedelta(days=30)

Valed_Date =new_date.strftime("%d\\%m\\%Y")



T2 ="BRP_410                                       "
Book_Cancel ="BRP_951                                       "
Last_Run = 0
Timer = 0
pause = 0
UnifiedCode1 = 0
CustodianCode = 0
QTY = 0
SQL_Query = 0
column_names = 0
UnifiedCode12_index = 0
QTY_index = 0
Action1_index = 0
Ticker_index = 0
Reuters_index = 0
SymbolCode = 0
UnifiedCode2 = 0
Symbol_Num = 0
Stock_Num = 0
comment = ''

#pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Thndr\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
pyautogui.PAUSE = 0.1
keyboard2 = Controller()

path = os.getcwd() + '\\data\\'


def pause_execution():
    global pause
    if pause == 0:
        pause = 1 
        answer = pyautogui.confirm("Do you want to Exit ?", "PAUSE", buttons=["Yes", "No"])
        if answer == "Yes":
            #pause = 0
            # Hide the windows or perform any other actions
            pause = 2
            ui.Status_label.setText("Operations have been cancelled..")
            QApplication.processEvents()
            return
        elif answer == "No":
            pause = 0

#keyboard.add_hotkey('ALT+ p', pause_execution)         #----------------add_hotkey
keyboard.add_hotkey('ALT+p', pause_execution)



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
            if click == 2 :
               pyautogui.doubleClick(x- offset, y)

            elif click == 1:
                pyautogui.leftClick(x-offset,y, interval=0.001)
            else :
                pass

        except Exception as e:
            time.sleep(0.1)
            
        else:
            return x , y
            break
    pyautogui.alert("Can't Find "+ img ,"Error")
    
            


stop_flag = threading.Event()

def detect_and_click():
    while not stop_flag.is_set() and pause == 0 :
        # Check if the image appears on the screen
        if pyautogui.locateOnScreen('Issue_Num.png') is not None :
            time.sleep(0.1)
            # Perform the click action on the image
            pyautogui.click('Issue_Num_Ok.png')
            #break
             

#image_detection_thread = threading.Thread(target=detect_and_click)
#image_detection_thread.start()




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
        pyautogui.alert("Can't Find "+ app_title ,"Error")

        return False



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
    pyautogui.alert("Can't Find "+ img1 ,"Error")

    return x ,y,img



def Booking_Cancellation():
    global pause
    global UnifiedCode2
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
       return error
    else:
       time.sleep(0.2)
       clickOn('Custodian_Code.png',70,2,10)
       time.sleep(0.2)
       keyboard2.type(CustodianCode)

       time.sleep(0.3)
       keyboard2.type('\t')  
       time.sleep(0.2)
       clickOn("The_Client.png",90,2,30)
       Change_language() 

       time.sleep(0.5)
       keyboard2.type(UnifiedCode2)
      
       keyboard2.type('\t')  

       time.sleep(0.5)
       clickOn("The_Stock.png",90,2,30)
       time.sleep(0.4)
       keyboard2.type(Symbol_Num)
       time.sleep(0.4)
       keyboard2.type('\t')  
       time.sleep(0.5)               
       keyboard2.type('31122000')
       time.sleep(0.5)
       clickOn('Search.png',0,1,30)
       time.sleep(2)
       
       try:
            x,y,img = Detect_img("Remain.png",0,0,10)
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
            print(Stock_Num)
            

            #Stock_Num = pyperclip.paste()

            if Stock_Num.strip().isdigit() and int(Stock_Num) > 0 :
                clickOn('Delete.png',0,1,40)
                time.sleep(0.2)
                x,y,img = Detect_img('Order_Is_Deleted.png','Cannot_delete.png','Delete_Comment.png',30)
                if img == 1:
                    time.sleep(0.3)
                    clickOn('Order_Is_Deleted_ok.png',0,1,20)
                    time.sleep(0.6)
                    error = 0
                    return error
                elif img == 2:
                    time.sleep(0.2)
                    clickOn('FRM_40401_Ok.png',0,1,8)
                    time.sleep(0.4)
                    error = 3
                    return error
                elif img == 3 :
                    time.sleep(0.2)    
                    clickOn('Delete_Comment.png',90,2,10)
                    keyboard2.type('cancel')
                    time.sleep(0.2)
                    clickOn('Ok.png',0,1,30)
                    time.sleep(0.4)
                    x,y,img = Detect_img('Cannot_delete.png','Delete_Done.png','FRM_40401.png',30)
                    if img == 1:
                        time.sleep(0.2)
                        clickOn('Cannot_delete_ok.png',0,1,8)
                        time.sleep(0.4)
                        clickOn('back.png',0,1,10)
                        time.sleep(0.4)
                        error = 3
                        return error

                    elif img == 2:
                        #clickOn('Delete_Done.png',0,1,40)
                        time.sleep(0.3)
                        clickOn('Delete_Done_ok.png',0,1,40)
                        time.sleep(0.5)
                        error = 0
                        return error
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
                            return error
                        elif img == 2:
                            time.sleep(0.35)
                            clickOn('Delete_Done_Ok.png',0,1,10)
                            time.sleep(0.2)
                            error = 0
                            return error
                        else:
                            error = 3
                            return error
                            
                if img == None:
                        error = 3
                        return error
            else:
                error = 2
                return error



def Booking_T2():
    global pause
    global UnifiedCode2
    global Symbol_Num
    global Stock_Num
    global SymbolCode
    global CustodianCode
    global Valed_Date
    global QTY

    while True:
        if pause == 0 :
            break
        elif pause == 2 :
            return
  
    if not WinActivate(T2,4):
       error = 1
       return error
    else:
        time.sleep(0.2)

        clickOn('SymbolCode.png',120,2,30)

        Change_language() 

        keyboard2.type(SymbolCode)
        clickOn('The_QTY.png',120,2,30)
        if QTY < Stock_Num :
            QTY = Stock_Num
        keyboard2.type(QTY)
        time.sleep(0.3)
        clickOn('CustodianCode.png',120,2,30)
        keyboard2.type(CustodianCode)
        time.sleep(0.3)
        clickOn('The_Client2.png',120,2,30)
        keyboard2.type(UnifiedCode2)
        time.sleep(0.3)
        pyautogui.typewrite('\t') 
        time.sleep(0.3)
       
        keyboard2.type(Valed_Date)
        pyautogui.typewrite('\t') 
        clickOn('Save.png',0,1,30)
        x,y,img = Detect_img ('Not_Enough_Balance.png','Qty_Is_Booked.png','not_listed.png',40)
        if img == 1: 
            clickOn('Delete_Done_Ok.png',0,1,20)
            error = 2
            return  error
        elif img == 2 :
            time.sleep(0.2)
            clickOn('Qty_Is_Booked_Ok.png',0,1,40)
            return 0
        elif img == 3 :
            time.sleep(0.2)
            clickOn('Cannot_delete_ok.png',0,1,10)
            time.sleep(0.2)
            error = 3
            return error




def Query_Google_Sync():
    global pause
    global UnifiedCode1
    global CustodianCode
    global QTY
    global Last_Run  
    global UnifiedCode12_index
    global QTY_index 
    global Action1_index 
    global Ticker_index 
    global Reuters_index 
    global SymbolCode
    global UnifiedCode2
    global Symbol_Num
    global Stock_Num
    global comment
    filename = 'Q_results.csv' 

 #------------------------------------------------------ google sheet ----------------------------------------------------

    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return

    Last_Run = time.time()


    # Set up authentication credentials
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_file = 'creds.json'
    google_jsonfile = 'json.json'

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
        #sheet_url = 'https://docs.google.com/spreadsheets/d/1Y4HTSNjjTUiqeifuiVZg63tvRlIYIbTTkDpnEPOssnw/edit#gid=0'
        sheet = client.open_by_url(sheet_url).worksheet('Orders')
    else:
        pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
        return




    # Find the last row with values
    last_row = len(sheet.col_values(1)) + 1


    # get the index of the sheet1_header
    sheet1_header = sheet.row_values(1)
    UnifiedCode1_index = sheet1_header.index('UnifiedCode')
    Action1_index = sheet1_header.index('Action')
    Eligibility_Index = sheet1_header.index('Eligibility')


    #Sheet2 = client.open_by_url(sheet_url).worksheet('Results')
    #sheet2_header = Sheet2.row_values(1)
    
    Sheet3 = client.open_by_url(sheet_url).worksheet('MCDR Mapping')
   
   



    ui.Status_label.setText("Sync Google Sheet..")
    QApplication.processEvents()
    #filename.close()

    

    # find the unresolved  issues rows
    sheet1_row = sheet.get_all_values()

    

    for i, row in enumerate(sheet1_row):               # the main loop that excute the boobking steps-----------------------
        
        if row[Action1_index] == '' and  row[Eligibility_Index]==1 and  i>0:             # Skip sheet1_header row
            UnifiedCode1 = row[UnifiedCode1_index]
            
        
            with open(filename, 'r', encoding='utf-16') as csvfile:
               reader = csv.reader(csvfile)
               rows = list(reader)


           # Update the sheet with the CSV data starting from the last row
            SQL_Query.replace("'", "")
            sheet2_header = Sheet2.row_values(1)
            if sheet2_header != column_names:
             # Insert the column names in the first row
             Sheet2.insert_rows([column_names], row = 1)
             Sheet2.insert_rows(rows, row = last_row2 + 1)

           else:
             Sheet2.insert_rows(rows, row = last_row2)

           sheet2_header = Sheet2.row_values(1)  
           sheet.update_cell(i+1, Action1_index+1, 'Query done') 

                           
    sheet3_header = Sheet3.row_values(1)                   
    
    QTY_index = sheet2_header.index('Qty')
    CustodianCode_index = sheet2_header.index('CustodianCode')
    Reuters_index = sheet3_header.index('Reuters')

    sheet2_row = Sheet2.get_all_values()
    Action2_index = sheet2_header.index('Action')
    Ticker_index = sheet2_header.index('ticker')
    UnifiedCode2_index = sheet2_header.index('UnifiedCode')
    SymbolCode_index = sheet2_header.index('SymbolCode')

    for i, row in enumerate(sheet2_row):                                # -------------------------(loop)---------------------

        if row[Action2_index] == '' and i > 0 :
            while  True:
                if pause == 0 :
                    break
                elif pause == 2 :
                    return

            SymbolCode = row[SymbolCode_index]
            UnifiedCode2 = row [UnifiedCode2_index]            
            cell = Sheet3.find(SymbolCode)

            if cell is not None:          
                row_num = cell.row
                row_sheet3 = Sheet3.row_values(row_num)
                Symbol_Num = row_sheet3[3]
                QTY = row[QTY_index]
                CustodianCode = row[CustodianCode_index]

                comment = ''   
                error = Booking_Cancellation()  
                if error    == 1 :
                    ui.Status_label.setText("Can't open MCDR System ..")
                    QApplication.processEvents()
                    break
                elif error == 0:
                    comment = Stock_Num + ' canceled'
                   #Sheet2.update_cell(i+1, Action2_index+1, Stock_Num) 
                elif error ==3 :
                    comment = 'Some Qty in Market'

                else:
                    comment = comment + '0 stock canceled '
                if pause == 2 :
                    return    
                #error = Booking_T2()     
                #if error == 1:
                #   ui.Status_label.setText(" Can't open MCDR System ..")
                #   QApplication.processEvents()
                #   break 
                #   #comment = comment + 'and not Enough Balance for' + QTY 
                #   #Sheet2.update_cell(i+1, Action2_index+1, 'Not Enough Balance for' + QTY) 
                #elif error == 0 :
                #    comment = comment + ' and ' + QTY + ' Booked'
                #   #Sheet2.update_cell(i+1, Action2_index+1,'Booked '+ QTY)
                #elif error == 2 :
                #    comment = comment + ' and not Enough Balance for ' + QTY
                #elif error == 3:
                #    comment = comment + ' Code Suspended  '    
            else :
                comment = 'can not find Symbol_Num '

            
                
            Sheet2.update_cell(i+1, Action2_index+1, comment)
            comment = '' 
            if pause == 2 :
                return     
    

    
#------------------------------------------------------------------------------------------------------------------------------------
           
# ---------------------------------------------------------------Azure SQL Database connection details

def Query():

    global SQL_Query
    global column_names
    global pause

    server = ui.Server_input.text()

    database = ui.Db_input.text()
    username = ui.User_N_input.text()
    password = ui.PW_input.text()
    global Timer
    Timer =  ui.Seconds_num.text()
    if ui.Seconds_num.text() == '' :
       error = 4
       return error

    if os.path.exists('Query.txt'):    
        try:
            with open('Query.txt', 'r') as f:
             SQL_Query = f.read()
             SQL_Query = SQL_Query.replace('--*', UnifiedCode1 )


        except Exception as e:
            error = 1
            return error
    else:
        pyautogui.alert("Can't Find Query file" ,"Error")
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
    cursor = conn.cursor()
    cursor.execute(SQL_Query)
    query_results = cursor.fetchall()



    # Write the query results to the CSV file with proper encoding
    filename = 'Q_results.csv'
    with open(filename, 'w', newline='', encoding='utf-16') as csvfile:
        writer = csv.writer(csvfile)
        # Get the column names from the cursor's description
        column_names = [column[0] for column in cursor.description]
        column_names.append("Action")


        # Write each row of the query results
        writer.writerows(query_results)

    # Close the database connection
    conn.close()

    Last_Run = time.time()




def Run_Function():
    sleep_duration = 0
    global Timer
    global Last_Run
    global pause
    
    while True:                                            #----------------------------------------

        #keyboard.add_hotkey('ALT+ p', pause_execution)
        

        Timer = ui.Seconds_num.text()
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
           break
        elif error == 3:
           ui.Status_label.setText("Can't Find MCDR")
           QApplication.processEvents()
           break
        elif error == 4:
           ui.Status_label.setText("Please enter the number of seconds..")
           QApplication.processEvents()
           return
        elif pause == 2:
            pause = 0
            break
        
        else:
            Current_time = time.time()
            dev = float ( Current_time - Last_Run )
            
            ui.Status_label.setText("wait for next run....")
            QApplication.processEvents()
            sleep_duration = float(Timer) - dev
            print(sleep_duration)
            if sleep_duration > 0:
                loop = 0
                while loop < sleep_duration:
                    loop += 0.2
                    #time.sleep(sleep_duration)
                    time.sleep(0.2)
                    if pause == 2:
                       pause = 0 
                       return



def check_pause():
    global pause 
    if pause == 0:
       pause_execution()
    else:
        pass

        
        



if __name__ == "__main__":
    #import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()

    ui = Win.Ui_MainWindow()
    ui.setupUi(MainWindow)
  

    ui.Server_input.setText("10.70.4.100")     #put value in server IP
    ui.Db_input.setText("master")     #put value in Database Name

    #ui.User_N_input.setText('f.shawky')
    #ui.PW_input.setText('gHB*7J/')
    #MainWindow.setWindowTitle("Auto Booking")

    ui.PW_input.setEchoMode(QtWidgets.QLineEdit.Password)
    ui.Status_label.setAlignment(QtCore.Qt.AlignCenter)
    ui.start_b.clicked.connect(Run_Function)
    ui.puase_b.clicked.connect(check_pause)


    MainWindow.show()
    sys.exit(app.exec_())


