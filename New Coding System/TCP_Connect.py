# name = input("Enter your name: ")
# print("Hello,", name)
# age = input("Enter your Age:")
# print("good you are in:",age)
# male = input("are you Male or not?:")
# print("good you are :",male)


import socket  
from pynput.mouse import  Controller  
import time  
import random  
import subprocess
autoit_script = './Data/Download_Codes_TCP.au3'


mouse = Controller()  

def human_move():  

    duration = 0.4
    for i in range (15):
        destination = (random.randint(200, 1400 - 1), random.randint(100, 700 - 1))
        start_position = mouse.position  
        distance_x = destination[0] - start_position[0]  
        distance_y = destination[1] - start_position[1]  

        steps = 100  
        pause_duration = duration / steps  

        for i in range(steps):  
            x = int(start_position[0] + (distance_x * (i / steps)) + random.uniform(-1, 1))  
            y = int(start_position[1] + (distance_y * (i / steps)) + random.uniform(-1, 1))  
            mouse.position = (x, y)  
            time.sleep(pause_duration)  

import threading  

# duration = 0.4
# for i in range (15):
#     destination = (random.randint(200, 1400 - 1), random.randint(100, 700 - 1))
# human_move()
Run_human_move = threading.Thread(target=human_move)
Run_human_move.start()
print('haaaaaaaaaaaaaaai')

exit()

import datetime

def TCP_listen():
    current_date = datetime.datetime.now()
    print(current_date)

        
    host = '127.0.0.1'  
    port = 65432        

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:  
        s.bind((host, port))  
        s.listen()  
        print(f"Listening on {host}:{port}")  
        subprocess.Popen(['./Data/autoit3',autoit_script], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        while True:  

            
            conn, addr = s.accept()  
            with conn:  
                print(f"Connected by {addr}")  
                data = conn.recv(1024) 
                
                if not data:  
                    break  

                message = data.decode('utf-8')  
                print(f"Received data: {message}")

                reply = 'done'
                # response_message = message  
                conn.sendall(reply.encode('utf-8'))
                # conn.sendall(response_message)
                time.sleep(2) 
                current_date2 = datetime.datetime.now()
                if ((current_date2 - current_date).total_seconds())> 1:

                    print((current_date2 - current_date).total_seconds())



TCP_listen()
