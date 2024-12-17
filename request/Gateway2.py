import socket
import requests
from bs4 import BeautifulSoup
import pyautogui

s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
s.connect(("8.8.8.8", 80))
My_IP = s.getsockname()[0]
print(My_IP)

def check_website_connection(url, interface_ip):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind((interface_ip,0))
        # s.connect((url,80))
    print(f"Connection to {url} successful using {interface_ip}.")

    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    plain_text = soup.get_text()
    pyautogui.alert(plain_text)

    with open('output.txt', 'w') as a:
        pass

check_website_connection("https://google.com", My_IP)