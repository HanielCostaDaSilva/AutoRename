import time
import pyautogui

while(True):
    print(pyautogui.position())
    pyautogui.scroll(-20)

    time.sleep(5)