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
    

pyautogui.alert(table_rows)
