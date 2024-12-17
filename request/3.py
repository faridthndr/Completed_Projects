import requests
from bs4 import BeautifulSoup
import csv
from datetime import datetime

with open('Url.txt', 'r') as f:
    url = f.read()

with open('OutPut.csv', 'r') as csvfile:
    csv_reader = csv.reader(csvfile)
    try:
        first_row = next(csv_reader)
        if len(first_row) >= 3:
            Last_Date = first_row[2]
    except StopIteration:
        Last_Date = '2023-01-01' 
response = requests.get(url)
response.encoding = 'utf-8'
soup = BeautifulSoup(response.content, 'html.parser')

table_rows = soup.find_all('tr')
with open('output.csv', 'w', newline='', encoding='utf-8') as csvfile:
    csv_writer = csv.writer(csvfile)
    for row in table_rows:
        table_data = row.find_all('td')
        if len(table_data) >= 4:
            cell1, cell2, cell3, cell4 = table_data
            Date = cell3.text.strip()

            try:
                Date_and_Time = datetime.strptime(Date, '%d/%m/%Y %H:%M:%S').isoformat()
            except ValueError:
                Date_and_Time = Date

            if Date_and_Time > Last_Date:
                News_No = cell1.text.strip()
                Isin_code = cell2.text.strip()
                News_Body = cell4.text.strip()
                row_data = [News_No, Isin_code, Date_and_Time, News_Body]
                csv_writer.writerow(row_data)
