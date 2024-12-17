import csv
import time
import pyautogui
import sys
import subprocess
import os
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtWidgets, QtGui, QtCore
import Main_window as win
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import keyboard


Order_Type_Index=0
Order_Type= 0
Market_Code_index=0
Market_Code=0
ISIN_Code_index=0
ISIN_Code=0
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
X_stream_Index = 0
X_stream = 0
X_stream_Comment = 0
Comment_indx = 0

All_Even_Odd = 1
pause =0
Last_Run=0  
Action1_index =0
comment=0
filename = 'Q_results.csv' 
autoit_script = "execute.au3"


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


def check_pause():
    global pause 
    if pause == 0:
       pause_execution()
    else:
        pass


keyboard.add_hotkey('ALT+p', check_pause)



def Query_Google_Sync():


    global Order_Type_Index
    global Order_Type
    global Market_Code_index
    global Market_Code
    global ISIN_Code_index
    global ISIN_Code
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
    global X_stream_Index
    global X_stream 
    global Comment_indx
        
    global pause
    global Last_Run
    global Action1_index 
    global Ticker_index 
    global X_stream_Comment
    global comment
    global autoit_script
    global filename 
   

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
    google_jsonfile = 'offline-orders.json'

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

      
    for i, row in enumerate(sheet1_row):               # the main loop that excute the boobking steps-----------------------

        
        if pause == 2 :
           return
                        
        if row[CustodianCode_index] !='' and row[Thndr_Code_Index] != '' and row[Price_Index] != '' and row[QTY_index] !=''  and row[ISIN_Code_index] != ''and row[Market_Code_index] != '' and row[Order_Type_Index] != '' and row[Comment_indx] != 'Comment' :

           Order_Type = row[Order_Type_Index]
           Market_Code = row[Market_Code_index]
           ISIN_Code = row[ISIN_Code_index]
           QTY = row[QTY_index]
           Price = row[Price_Index]
           Thndr_Code = row[Thndr_Code_Index]
           CustodianCode = row[CustodianCode_index]

           
           if  row[Order_Type_Index]=='Sell' :

               if row[X_stream_Index] == '' and ( row[Offline_Booking_indx] !=''   or (row[CustodianCode_index] =='4625' or row[CustodianCode_index] == '4503')): 

                   
                   x_stream()

                   sheet.update_cell(i+1, X_stream_Index + 1, X_stream_Comment)

               if (row[CustodianCode_index] == '4625' or row[CustodianCode_index] =='4503' ) and row [X_stream_Index] == 'Qty Not Booked' and row[Offline_Booking_indx] != '' and row[Comment_indx] == '' :

                   x_stream()
                   sheet.update_cell(i+1, X_stream_Index + 1, X_stream_Comment)
                   sheet.update_cell(i+1, Comment_indx + 1 , 'Tried twice')
                   # u.close()
           if  row[Order_Type_Index]=='Buy' and  row[X_stream_Index] =='' and row[Purchase_Power_Indix] !=''  :
                    x_stream()

    ui.Status_label.setText("Done..")
    QApplication.processEvents()




def x_stream():
    global Order_Type
    global Market_Code
    global ISIN_Code
    global QTY
    global Price
    global Thndr_Code
    global CustodianCode
    global autoit_script
    global X_stream_Comment


    with open("Input_var.txt", "w") as f:
      f.write( Order_Type + "\n")        
      f.write( Market_Code + "\n")
      f.write( ISIN_Code + "\n")
      f.write( QTY + "\n")
      f.write( Price +"\n")
      f.write( Thndr_Code +"\n")
      f.write( CustodianCode + "\n")

    f.close()

    if os.path.exists("OutPut.txt"):
       os.remove("OutPut.txt")
    process = subprocess.Popen(['autoit3',autoit_script], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    output, error = process.communicate()
    process.communicate()
    time.sleep(0.4)
    with open("OutPut.txt", "r") as u:
        X_stream_Comment = u.readline()
    u.close()




def check_box(box_num):
    global All_Even_Odd
    if  box_num == 1 and   ui.All_Rows.isChecked() :  
        ui.Even_Rows.setChecked(False) 
        ui.Odd_Rows.setChecked(False)
        All_Even_Odd = 1
    elif box_num == 2 and  ui.Even_Rows.isChecked() :
        ui.All_Rows.setChecked(False) 
        ui.Odd_Rows.setChecked(False) 
        All_Even_Odd = 2
    elif box_num == 3 and  ui.Odd_Rows.isChecked() :
        ui.All_Rows.setChecked(False) 
        ui.Even_Rows.setChecked(False) 
        All_Even_Odd = 3


def Run_Fun():
    global pause
    while  True:
        Query_Google_Sync()
        
        if pause == 2 :
            return

        for _ in range(150):        
            

            if pause == 2 :
                return
            time.sleep(0.1)

      





if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = win.Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.start_b.clicked.connect(Run_Fun)
    ui.Pause_b.clicked.connect(pause_execution)
    ui.All_Rows.setChecked(True)
    ui.All_Rows.stateChanged.connect(lambda: check_box(1))
    ui.Even_Rows.stateChanged.connect(lambda: check_box(2))
    ui.Odd_Rows.stateChanged.connect(lambda: check_box(3))
    

    MainWindow.show()
    sys.exit(app.exec_())
