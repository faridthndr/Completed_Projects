import requests
from bs4 import BeautifulSoup
import pyautogui
from bs4 import BeautifulSoup



url = 'https://www.alrab7on.com/' 
response = requests.get(url)

soup = BeautifulSoup(response.text, 'html.parser')

plain_text = soup.get_text()

pyautogui.alert(plain_text)

with open('output.txt' ,'w') as a:
	pass
	# a.write(plain_text)