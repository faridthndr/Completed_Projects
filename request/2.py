import requests
from bs4 import BeautifulSoup
import pyautogui

text = ''

with open('Url.txt', 'r') as f:
    url = f.read()

response = requests.get(url)
soup = BeautifulSoup(response.text, 'html.parser')

table_rows = soup.find_all('tr')
for row in table_rows:
    table_data = row.find_all('td')
    if len(table_data) >= 4:
        cell1, cell2, cell3, cell4 = table_data
        text += f"{cell1.text}\n{cell2.text}\n{cell3.text}\n{cell4.text}\n"

pyautogui.alert(text)
