import pyautogui
import time 

#stopping program contingency
pyautogui.FAILSAFE = True

#number of properties we are working with
num_prop = 500

for i in range(num_prop):
    pyautogui.click(x=1283, y=26) #click on excel sheet
    pyautogui.press("down") #scroll down (press down key)
    pyautogui.hotkey('ctrl', 'c') #copy (ctrl c)
    time.sleep(0.1)
    pyautogui.moveTo(x=3456, y=42) #put cursor in file explorer
    pyautogui.click() 
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v') #paste (ctrl v)
    pyautogui.press("enter") #save it
    time.sleep(1.5)
