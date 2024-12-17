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




#Valed_Date =new_date.strftime("%d\\%m\\%Y")



BK_win ="BOP_150_1                                       "
Timer = 0
pause = 0
UnifiedCode = 0
Expirydate = 0
UnifiedCode_index = 0
Action_index = 0
comment = ''
app_title_win = 'Users Registration'

#pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Thndr\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
pyautogui.PAUSE = 0.1
keyboard2 = Controller()

current_date = current_date = datetime.datetime.now()
Today_date = current_date.strftime("%d-%m-%Y")

new_date = current_date + datetime.timedelta(days=730)
two_years_later = new_date.strftime("%d-%m-%Y")



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
    
            


#stop_flag = threading.Event()

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


def Registration():
    global BK_win 
    global Timer
    global pause 
    global UnifiedCode 
    global Expirydate 
    global UnifiedCode_index 
    global Action_index 
    global comment
    global two_years_later
    global Today_date

    comment = ''

    current_date = current_date = datetime.datetime.now()
    Today_date = current_date.strftime("%d-%m-%Y")

    new_date = current_date + datetime.timedelta(days=730)
    two_years_later = new_date.strftime("%d-%m-%Y")



    while True:
        if pause == 0 :
            break
        elif pause == 2 :
            return
  
  
    if not WinActivate(BK_win,4):
       error = 11
       return error
    else:
        time.sleep(0.1)

        clickOn('User_Code.png',80,1,30)

        Change_language() 

        keyboard2.type(UnifiedCode)
        keyboard.press("tab")
        x,y,img = Detect_img('User_Not_reg.png','User_Already_Reg.png','Need_to_Confirm.png',30)
        if img == 1:
            clickOn('ok.png',0,1,20)
            
            x,y,img = Detect_img('User_Reg.png','User_Reg2.png',0,30)
            if img == 1:    
                clickOn('User_Reg.png',0,1,30)
            if img == 2:
                clickOn('User_Reg2.png',0,1,30)

            time.sleep(0.1)
            clickOn('Reg_Date.png',0,1,50)
            keyboard2.type(Today_date)
            clickOn('Next_Reg_Date.png',80,1,30)
            keyboard2.type(two_years_later)
            time.sleep(0.15)
            clickOn('Save.png',0,1,30)
            time.sleep(0.15)
            x,y,img = Detect_img('Confirm_Msg.png',0,0,30)
            if img == 1:
                clickOn('ok.png',0,1,40)
                time.sleep(0.1)
                x,y,img = Detect_img('Reg_Don_Msg.png',0,0,30)
                if img == 1 :
                    clickOn('ok.png',0,1,20)
                    comment = 'Done'
                    return comment


        if img == 2 or img == 3 :
            clickOn('known.png',0,1,30)
            time.sleep(0.2)
            clickOn('clear.png',0,1,30)
            if  img == 2:
                comment = 'Already exist'
            else:
                comment = 'Need to Confirm'    
                pass
            x,y,img = Detect_img('Empty_Data.png',0,0,30)
            if img == 1:

                return comment
            else:
                comment = 4
                return comment


            




def Query_Google_Sync():
    global BK_win 
    global Timer
    global pause 
    global UnifiedCode 
    global Expirydate 
    global UnifiedCode_index 
    global Action_index 
    global comment

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
        sheet = client.open_by_url(sheet_url).worksheet('New_Users_Reg')
    else:
        pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
        error = 3
        return error 


    # Find the last row with values
    #last_row = len(sheet.col_values(1)) + 1                  #-*****************


    # get the index of the sheet1_header
    sheet1_header = sheet.row_values(1)

    UnifiedCode_index = sheet1_header.index('UnifiedCode')
    Action_index = sheet1_header.index('Action')
    

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

        if sheet1_row[i][Action_index] == '' and sheet1_row[i][UnifiedCode_index] !='' :

           UnifiedCode = sheet1_row[i][UnifiedCode_index]
           
           comment = Registration()

           if comment != 4 :
               sheet.update_cell(i+1, Action_index + 1, comment)
               comment = '' 
           else:
               error = 4
               return error

           if pause == 2 :
              return
    WinActivate(app_title_win,5)            






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
    MainWindow.setWindowTitle(app_title_win)

    ui.Status_label.setAlignment(QtCore.Qt.AlignCenter)
    ui.start_b.clicked.connect(Run_Function)
    ui.Pause_b.clicked.connect(pause_execution)

    MainWindow.show()
    sys.exit(app.exec_())


