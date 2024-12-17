from oauth2client.service_account import ServiceAccountCredentials
import gspread
import os
from tkinter import messagebox
import csv

csv_file_path = 'All_Users_data.csv'



scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds_file = 'creds.json'
google_jsonfile = 'coding-new-users.json'

creds = None
if os.path.exists(creds_file):         
    with open('creds.json', 'r') as f:  
        creds_json = f.read()
        creds = ServiceAccountCredentials.from_json(creds_json)     
else:
    if not os.path.exists(google_jsonfile):     
        error = 3
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name(google_jsonfile, scope)   
        with open(creds_file , 'w') as f:
            f.write(creds.to_json())

client = gspread.authorize(creds)


if os.path.exists('Sheet_Url2.txt'): 
        with open('Sheet_Url2.txt', 'r') as u:   
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
# google_rows = sheet.get_all_values()
google_rows = CSV_sheet.get_all_values()

# ID_Num_Col = [row[14] for row in google_rows]
ID_Num_Col = [row[9] for row in google_rows]


# if not os.path.isfile(csv_file_path):  
#     print("not found ")
#     exit()
# else:
#     print("found")
#     column_o_values = [row[14] for row in google_rows]  

#     try:
#         with open (csv_file_path,'r',encoding='utf-8') as csvfile:
#             row_data =  csv.reader(csvfile) 
#             rows = list(row_data)[1:]

#         for row in rows:
#              value_to_compare = row[9]
#              values_to_add = row[18:26]
#              if value_to_compare in column_o_values:
#                   row_index = column_o_values.index(value_to_compare) + 1
#                   update_row = google_rows[row_index-1][23:31]
#                   update_row[:] = values_to_add  
                
#                   sheet.update(range_name=f'X{row_index}', values=[update_row])  

#     except Exception as e:
#         messagebox.showinfo(e)


import socket  

host = '127.0.0.1'   
port = 65432         

with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:  
    s.bind((host, port))  
    s.listen()  
    print(f"Listening on {host}:{port}")  
    
    while True: 
        try: 
            conn, addr = s.accept()  
            with conn:  
                print(f"Connected by {addr}")  
                data = conn.recv(1024) 
                if not data:  
                    break  
                message = data.decode('utf-8')
                message =  message.split(',')                 

                print(f"Received data: {message}")
                # value_to_compare = message[0]
                value_to_compare = message[9]

                # values_to_add = message[1:]

                values_to_add = message[18:26]
                if value_to_compare in ID_Num_Col:
                    row_index = ID_Num_Col.index(value_to_compare) +1
                    update_row = google_rows[row_index-1][18:26]
                    update_row[:] = values_to_add  
                    CSV_sheet.update(range_name=f'S{row_index}', values=[update_row])              
                    reply = 'Done'
                else:
                    reply = 'failed'
                    # response_message = message  
                conn.sendall(reply.encode('utf-8'))
                    # conn.sendall(response_message) 
        except Exception as e:
            print(e)
            reply = 'failed'
            conn.sendall(reply.encode('utf-8'))
            # conn.close()    # if we need to close the connection


            


