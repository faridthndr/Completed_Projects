import os 
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pyautogui
import json
import tkinter as tk
from tkinter import messagebox
import requests
import pyodbc
import decimal
import tkinter as tk  
from tkinter import simpledialog  
import sys

root = tk.Tk()  
root.geometry("400x200")  
root.title("User Credentials") 
root.withdraw() 


username = simpledialog.askstring("User Input", "Please enter your username:")  
password = simpledialog.askstring("User Input", "Please enter your password:", show='*')
server = '10.70.4.100'
database = 'master'
# username = 'farid.shawky@thndr.app'
# password = '8mrM7dZF@5LP'
if username == '' or  password == '' :
 	pyautogui.alert("Error in username or password ","Error")
 	sys.exit()
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
      
      error = 3
      # return error
   else:
      creds = ServiceAccountCredentials.from_json_keyfile_name(google_jsonfile, scope)   # if creds.json file not exist we get the creds by googlekey.json file as usual
      with open(creds_file , 'w') as f:
        f.write(creds.to_json())


client = gspread.authorize(creds)


if os.path.exists('Sheet_Url.txt'): 
    with open('Sheet_Url.txt', 'r') as u:   #--------------read URL for google sheet
        sheet_url = u.read()

    #Access the Google Sheet
    sheet = client.open_by_url(sheet_url).worksheet('Close Prices')
else:
    pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
    # return


try:
	with open('Query.txt', 'r') as f:
	 SQL_Query = f.read()
	   
except Exception as e:
    error = 1
    # return error

# Establish connection to Azure SQL Database
try:
    conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    conn = pyodbc.connect(conn_str)
    # Execute the SQL query and fetch the results
    conn.commit()
    cursor = conn.cursor()
    cursor.execute(SQL_Query)
    query_results = cursor.fetchall()

    new_results = [[float(value) if isinstance(value, decimal.Decimal) else value for value in row] for row in query_results]  


	# new_results = [list(row) for row in query_results] 
    sheet.update(values= new_results, range_name='A2:E')
    pyautogui.alert("Query Done ","Done")

except Exception as e:
    pyautogui.alert(f"Error: {e}",)
    # return error




    