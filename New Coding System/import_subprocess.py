
import subprocess
import tkinter as tk
from tkinter import messagebox


data = "mohamed`n47`nMale"  
command = f'$data = "{data}"; (echo "$data") | python "D:\\brokerage-scripts\\test.py"'  

try:  
    result = subprocess.run(["powershell.exe", "-Command", command], capture_output=True, text=True)  
    if result.returncode ==0:  
         print("results:")  
         print(result.stdout)
         root = tk.Tk()
         root.withdraw()
         messagebox.showinfo("Done", result.stdout)

    else:  
         print("Error:")  
         print(result.stderr)  
except Exception as e:  
 print(f"error {e}")  
