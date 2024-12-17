import pyautogui
from PIL import Image
import pytesseract
from pynput import mouse
import tkinter as tk
from tkinter import messagebox
import sys
import os


raise ValueError("Invalid input")
exit()


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
       
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Thndr\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = resource_path("tesseract/tesseract.exe")
start_x = start_y = end_x = end_y = None

def on_click(x, y, button, pressed):
    global start_x, start_y, end_x, end_y
    if pressed:
        if start_x is None and start_y is None:
            start_x, start_y = x, y
            print(f"First click at ({start_x}, {start_y})")
    else:
        if start_x is not None and start_y is not None:
            end_x, end_y = x, y
            print(f"Second click at ({end_x}, {end_y})")
            return False  # Stop listener

# Collect events until released
    
root = tk.Tk()
root.withdraw()  # Hide the main window

tk.messagebox.showinfo("Select Area", "Please click and drag to select the area to capture.")
root.update()

with mouse.Listener(on_click=on_click) as listener:
    listener.join()



if None not in (start_x, start_y, end_x, end_y):
    left, top = min(start_x, end_x), min(start_y, end_y)
    width, height = abs(end_x - start_x), abs(end_y - start_y)

    # Capture screenshot of the selected region
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot.save("selected_area.png")

    # Convert screenshot to text using OCR
    text = pytesseract.image_to_string(screenshot, lang='eng')
    pyautogui.alert(text)
else:
    print("Failed to capture mouse clicks.")
