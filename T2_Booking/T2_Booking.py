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
Expirydate = 0
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



def Detect_img(img1,img2,img3,img4,img5, time_out):
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
        if img4 != 0 : 
            try:
                x,y =  pyautogui.locateCenterOnScreen(path+img4, confidence=0.8)
            except Exception as e:
                time.sleep(0.1)
                x, y = None, None
                img = None
            else:
                img = 4
                return x, y , img
                break 
        if img5 != 0 : 
            try:
                x,y =  pyautogui.locateCenterOnScreen(path+img5, confidence=0.8)
            except Exception as e:
                time.sleep(0.1)
                x, y = None, None
                img = None
            else:
                img = 5
                return x, y , img
                break
                 

    x, y = None, None
    img = None
    pyautogui.alert("Can't Find "+ img1 ,"Error")

    return x ,y,img


def Booking_T2():
    global pause
    global UnifiedCode2
    global UnifiedCode1
    global SymbolCode
    global CustodianCode
    global QTY
    global Expirydate
    global T2
    global Symbol_Num
    global Stock_Num
    
    global Valed_Date

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

        Change_language() 

        keyboard2.type(SymbolCode)
        clickOn('The_QTY.png',120,2,30)
        #if QTY < Stock_Num :
        #    QTY = Stock_Num
        keyboard2.type(QTY)
        time.sleep(0.3)
        clickOn('CustodianCode.png',120,2,30)
        keyboard2.type(CustodianCode)
        time.sleep(0.3)
        clickOn('The_Client2.png',120,2,30)
        keyboard2.type(UnifiedCode1)
        time.sleep(0.3)
        pyautogui.typewrite('\t') 
        time.sleep(0.3)
       
        keyboard2.type(Expirydate)
        pyautogui.typewrite('\t') 
        clickOn('Save.png',0,1,30)
        x,y,img = Detect_img ('Not_Enough_Balance.png','Qty_Is_Booked.png','not_listed.png','booking_done_date.png','clinet_not_correct.png',40)
        if img == 1: 
            clickOn('Delete_Done_Ok.png',0,1,20)
            error = 111
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
            error = 112
            return error
        elif img == 4 :
            time.sleep(0.2)
            clickOn('Qty_Is_Booked_Ok.png',0,1,10)
            time.sleep(0.2)
            error = 113
            return error
        elif img == 5 :
            time.sleep(0.2)
            clickOn('Cannot_delete_ok.png',0,1,10)
            time.sleep(0.2)
            error = 114
            return error     
    



def Query_Google_Sync():
    global pause
    global UnifiedCode1
    global CustodianCode
    global QTY
    global Expirydate
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


    if ui.First_Row.text()=='' or ui.Last_Row.text()=='' or int(ui.First_Row.text()) <2 :
       pyautogui.alert("Please Enter first and last row " ,"Error")
       error = 1
       return error 


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
          error = 2
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
        sheet = client.open_by_url(sheet_url).worksheet('T2_Qty')
    else:
        pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
        error = 3
        return error 


    # Find the last row with values
    #last_row = len(sheet.col_values(1)) + 1                  #-*****************


    # get the index of the sheet1_header
    sheet1_header = sheet.row_values(1)
    UnifiedCode1_index = sheet1_header.index('UnifiedCode')
    SymbolCode_index = sheet1_header.index('SymbolCode')
    CustodianCode_index = sheet1_header.index('CustodianCode')
    QTY_index = sheet1_header.index('Qty')
    Expirydate_index = sheet1_header.index('Expiry_date')
    Result_index = sheet1_header.index('Result')

    First_Row = int(ui.First_Row.text())
    Last_Row =  int (ui.Last_Row.text())



    ui.Status_label.setText("Sync Google Sheet..")
    QApplication.processEvents()
    #filename.close()

    

    # find the unresolved  issues rows
    sheet1_row = sheet.get_all_values()
    count_row = len(sheet.col_values(1)) 
    if Last_Row > count_row:
       pyautogui.alert("Last Row number is greater than sheet rows") 
       error =1
       return error

    

    for i in range(First_Row - 1 , Last_Row ):                                                   # the main loop that excute the boobking steps-----------------------

        if sheet1_row[i][Result_index] == '' and sheet1_row[i][UnifiedCode1_index] !='' and sheet1_row[i][SymbolCode_index] != '' and sheet1_row[i][QTY_index] != '' and sheet1_row[i][Expirydate_index] !='':

           UnifiedCode1 = sheet1_row[i][UnifiedCode1_index]
           SymbolCode = sheet1_row[i][SymbolCode_index]
           CustodianCode = sheet1_row[i][CustodianCode_index]
           QTY = sheet1_row[i][QTY_index]
           Expirydate = sheet1_row[i][Expirydate_index]


           error = Booking_T2()

           if error == 111:
               comment = 'Not Enough Balance' 

           elif error == 0 :
                comment =  QTY + ' Booked'

           elif error == 112 :

                comment = 'Client Not listed'

           elif error == 113:
                comment ='Client Need a Contract  '   

           elif error == 114:
                comment ='The Client Not Correct '    

           sheet.update_cell(i+1, Result_index + 1, comment)
           comment = '' 
           if pause == 2 :
              return             






def Run_Function():
    sleep_duration = 0
    global Timer
    global Last_Run
    global pause
    
                                           #----------------------------------------
    
    error = Query_Google_Sync()
    if error == 1:
       ui.Status_label.setText("error in rows numbers..")
       QApplication.processEvents()
       return
    elif error == 2:
       ui.Status_label.setText("Json file does not exist.")
       QApplication.processEvents()
       
    elif error == 3:
       ui.Status_label.setText("Can't Find Sheet_Url.txt")
       QApplication.processEvents()
       
    elif error == 11 :
       ui.Status_label.setText("Can't open MCDR System.")
       QApplication.processEvents()
       return
    else:
     ui.Status_label.setText("Done.")
     QApplication.processEvents()

    if pause == 2:
        pause = 0
    
    


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
    ui.Status_label.setAlignment(QtCore.Qt.AlignCenter)
    ui.start_b.clicked.connect(Run_Function)
    ui.Pause_b.clicked.connect(pause_execution)

    MainWindow.show()
    sys.exit(app.exec_())


