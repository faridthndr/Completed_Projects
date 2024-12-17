import requests
import os
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import re



def Download(pdf_link,File_Name):
    Folder_Name = 'Downloaded_files'
    pattern = r'https://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)/'
    match = re.search(pattern, pdf_link)
    if match:
        file_id =  match.group(1)
    else:
        return None
        
    if not os.path.exists(Folder_Name):
        os.makedirs(Folder_Name)
    
    output = os.path.join(Folder_Name, File_Name)

    # file_id = '17VTB2IBtGwDFss-fRvmNQckWkWrS1SRQ'
    # file_id = 'https://drive.google.com/file/d/17VTB2IBtGwDFss-fRvmNQckWkWrS1SRQ/view'

    url = f'https://drive.google.com/uc?export=download&id={file_id}'

    # output = File_Name

    response = requests.get(url, stream=True)

    if response.status_code == 200:
        with open(output, 'wb') as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)
        print(f'File downloaded successfully as {output}')
    else:
        print(f'Failed to download file. Status code: {response.status_code}')


#-------------------------------------------------------------------------------------------------------#
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
    sheet = client.open_by_url(sheet_url).worksheet('Download_Links')
else:
    pyautogui.alert("Can't Find Sheet_Url.txt" ,"Error")
    # return

# Find the last row with values
last_row = len(sheet.col_values(1)) + 1

# get the index of the header
header =sheet.row_values(1)
Link_indx = header.index('PDF_Links')
File_Name_index = header.index('File_Name')
google_rows = sheet.get_all_values()

for i, row in enumerate(google_rows):   
    if i > 0:

        pdf_link = row[Link_indx]   
        output_file_name = row[File_Name_index] + ".pdf"
        Download(pdf_link,output_file_name)


#----------------------------------------------------------------------------------------------------------------------------------#

