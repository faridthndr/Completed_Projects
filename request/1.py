import requests
from bs4 import BeautifulSoup

text = ''

with open('Url.txt', 'r') as f:
  url = f.read()

response = requests.get(url)
response.encoding = 'utf-8'  

soup = BeautifulSoup(response.content, 'html.parser')

table_rows = soup.find_all('tr')
for row in table_rows:
  table_data = row.find_all('td')
  if len(table_data) >= 4:
    cell1, cell2, cell3, cell4 = table_data
    text1 = cell1.text.strip()
    text2 = cell2.text.strip()
    text3 = cell3.text.strip()
    text4 = cell4.text.strip()

    text += f"{text1}\n{text2}\n{text3}\n{text4}\n"

with open('output.txt', 'w', encoding='utf-8') as a:
  a.write(text)


