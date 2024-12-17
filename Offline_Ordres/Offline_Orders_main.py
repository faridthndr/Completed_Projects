
import time
import autoit
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
MCDR_Code_Index = 0
MCDR_Code = 0
Uni_Code_Index = 0
Uni_Code = 0



All_Even_Odd = 1
pause =0
Last_Run=0  
Action1_index =0
comment=0
filename = 'Q_results.csv' 
autoit_script = "execute.au3"
T0 ="BRP_800                                       "
T1 ="BRP_702                                       "

    

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

        

def WinActivate(app_title, time_out):
    global pause
   
    for _ in range(time_out):

        try:
            app = autoit.win_activate("[CLASS:{}]".format(app_title))
            time.sleep(0.2)
            return True
        except Exception as e:
            #print(f"App window with title '{app_title}' not found or could not be activated.")
            
            
            time.sleep(1)
            
    
    pyautogui.alert("Can't Find "+ app_title ,"Error")
    return False





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





def Change_language(hex_code):
    if SendMessage( GetForegroundWindow(), WM_INPUTLANGCHANGEREQUEST, 0, hex_code) == 0:
       SendMessage( GetForegroundWindow(), WM_INPUTLANGCHANGEREQUEST, 0, hex_code) 




keyboard.add_hotkey('ALT+p', check_pause)




def T0_call():
    global MCDR_Code_Index 
    global MCDR_Code
    global Uni_Code
    global ISIN_Code
    global pause

    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return

    T0_win = autoit.win_activate(T0)
        
    if not autoit.win_wait_active(T0,5):
       #pyautogui.alert("Can't Find T+0 ","Error")
       Stock_Num = 'not_found'
       return Stock_Num
    else:
       Change_language(0x4090409) 
       time.sleep(0.1)
       clickOn("client.png",90,2,5)
       time.sleep(0.1)
       pyautogui.typewrite(Uni_Code)
       clickOn("stock.png",90,2,5)
       pyautogui.typewrite(ISIN_Code)
       time.sleep(0.2)
       clickOn("search.png",0,1,5)
       time.sleep(0.25)

       clickOn("Issue_Num_Ok.png",0,1,5)
       time.sleep(0.4)
       clickOn("search.png",0,1,10)

       try:
          
            x,y,img = Detect_img("available_to_booked.png",0,0,10)
       except Exception as e:
            Stock_Num = 'not_found'
            return Stock_Num
       else:
            pyautogui.doubleClick(x, y+35)
            # Empty the clipboard
            pyperclip.copy('')
             # Simulate pressing Ctrl + C to copy the text
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.3)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.3)


            # Retrieve the copied text from the clipboard
            Stock_Num = pyperclip.paste()

            if Stock_Num.strip().isdigit() and int(Stock_Num) > 0:
               clickOn('record.png', 0, 1,20)

               x,y,img = Detect_img("confirm_Win.png",0,0,10)
               if img == None :

                   Stock_Num = 'not_found'
                   return Stock_Num
               else:
                   time.sleep(0.15)
                   clickOn("booked_Button.png", 0, 1,10)

                   x,y,img = Detect_img("Max_QTY.png","Done_MSG.png",0,20)
                   if img == 1 :
                       time.sleep(0.3)
                       clickOn("done.png",0,1,30) 
                       time.sleep(0.3)
                       clickOn("back.png",0,1,30)    
                       Stock_Num = 'MAx_QTY'
                       return Stock_Num
                   elif img == 2:
                       time.sleep(0.3)
                       clickOn("done.png", 0, 1,10)
                       return Stock_Num 
                       
                   else:   
                          Stock_Num = 'Error'
                          return Stock_Num
                      


def T1_call(Symbol_Num,UNI_Code):
    global pause
    while  True:
        if pause == 0 :
            break
        elif pause == 2 :
            return


    if not WinActivate(T1,3):
       Stock_Num = 'not_found'
       return Stock_Num
    else:
       Change_language(0x4090409) 
       time.sleep(0.1)
       clickOn("client.png",90,2,20)
       time.sleep(0.1)
       pyautogui.typewrite(UNI_Code)
       clickOn("stock.png",90,2,20)
       pyautogui.typewrite(Symbol_Num)
       time.sleep(0.2)
       clickOn("search.png",0,1,20)
       time.sleep(0.2)
       


       #clickOn("Issue_Num_Ok.png",0,1,10)
       #time.sleep(0.2)
       #clickOn("Issue_Num_Ok.png",0,1,15)
       time.sleep(0.2)  
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
    global X_stream_Index
    global X_stream 
        
    global pause
    global Last_Run
    global Action1_index 
    global Ticker_index 
    global comment
    global autoit_script
    global filename
    global MCDR_Code_Index 
    global MCDR_Code
    global Uni_Code_Index 
    global Uni_Code

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
    google_jsonfile = 'googlekey.json'

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
    X_stream_Index = sheet1_header.index('X_stream')
    MCDR_Code_Index = sheet1_header.index('MCDR Code')
    Uni_Code_Index = sheet1_header.index('Uni Code')


    
    

    ui.Status_label.setText("Sync Google Sheet..")
    QApplication.processEvents()
    #filename.close()

    

    sheet1_row = sheet.get_all_values()
    with open("Input_var.csv", 'w', newline='', encoding='utf-16') as csvfile:                
         writer = csv.writer(csvfile)
         #writer.writerows(sheet1_row)

    return  
    for i, row in enumerate(sheet1_row):               # the main loop that excute the boobking steps-----------------------
                    
        if sheet1_row[i][X_stream_Index] == '' and sheet1_row[i][CustodianCode_index] !='' and sheet1_row[i][Thndr_Code_Index] != '' and sheet1_row[i][Price_Index] != '' and sheet1_row[i][QTY_index] !=''  and sheet1_row[i][ISIN_Code_index] != ''and sheet1_row[i][Market_Code_index] != '' and sheet1_row[i][Order_Type_Index] != '' :

           

           Order_Type = sheet1_row[i][Order_Type_Index]
           Market_Code = sheet1_row[i][Market_Code_index]
           ISIN_Code = sheet1_row[i][ISIN_Code_index]
           QTY = sheet1_row[i][QTY_index]
           Price = sheet1_row[i][Price_Index]
           Thndr_Code = sheet1_row[i][Thndr_Code_Index]
           CustodianCode = sheet1_row[i][CustodianCode_index]
           MCDR_Code = sheet1_row[i][MCDR_Code_Index]
           Uni_Code = sheet1_row[i][Uni_Code_Index]

           if  sheet1_row[i][Order_Type_Index] !='Sell' and  sheet1_row[i][CustodianCode_index] !='46256':

                Stock_Num =  T0_call()


           writer.writerows(sheet1_row[i])

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
           os.system(autoit_script)

           output, error = process.communicate()
           time.sleep(0.2)
           with open("OutPut.txt", "r") as u:
                comment = u.readline()

           sheet.update_cell(i+1, X_stream_Index + 1, comment)

    ui.Status_label.setText("Done..")
    QApplication.processEvents()



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




if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = win.Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.start_b.clicked.connect(Query_Google_Sync)
    ui.Pause_b.clicked.connect(pause_execution)
    ui.All_Rows.setChecked(True)
    ui.All_Rows.stateChanged.connect(lambda: check_box(1))
    ui.Even_Rows.stateChanged.connect(lambda: check_box(2))
    ui.Odd_Rows.stateChanged.connect(lambda: check_box(3))
    

    MainWindow.show()
    sys.exit(app.exec_())
