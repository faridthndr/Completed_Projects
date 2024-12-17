import PyPDF2
from PyPDF2 import PdfReader
import requests
import os
import win32api
import win32print
import pyautogui
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from googleapiclient.discovery import build
from io import BytesIO

def Google_oauth2():
    global drive_service
    global folder_id

    folder_id = '1YzDiAOHeTVe6A4NbjvZQJx-JnmYw-NWN'
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_file = 'creds.json'
    google_jsonfile = 'googlekey.json'

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
            with open(creds_file, 'w') as f:
                f.write(creds.to_json())

    client = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)

    if os.path.exists('Sheet_Url.txt'):
        with open('Sheet_Url.txt', 'r') as u:
            sheet_url = u.read()

        # Access the Google Sheet
        Sheet = client.open_by_url(sheet_url).worksheet('Daily_News')
    else:
        pyautogui.alert("Can't Find Sheet_Url.txt", "Error")

    file_link = 'https://drive.google.com/file/d/17VTB2IBtGwDFss-fRvmNQckWkWrS1SRQ/edit'
    response = requests.get(file_link)
    pdf_file = BytesIO(response.content)

    try:
        pdf_reader = PdfReader(pdf_file)
        number_of_pages = len(pdf_reader.pages)  # Get number of pages using PdfReader
    except PyPDF2.errors.PdfReadError as e:
        print(f"Error reading PDF file: {e}")
        print("The PDF file may be corrupted or in an unsupported format.")
        # Handle the error, e.g., display an error message, try a different PDF library, or take alternative actions
        return

    temp_pdf_path = "temp_file.pdf"
    with open(temp_pdf_path, "wb") as f:
        f.write(response.content)

    printer_name = win32print.GetDefaultPrinter()

    win32api.ShellExecute(
        0,
        "print",
        temp_pdf_path,
        None,
        ".",
        0
    )

    os.remove(temp_pdf_path)

# Run the function to execute
Google_oauth2()