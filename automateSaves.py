import pyautogui
import time 

#stopping program
pyautogui.FAILSAFE = True
#initial selection of the yardi window
pyautogui.click(x=4905, y=153) 

#number of properties we are working with
num_prop = 255

for i in range(num_prop):
    pyautogui.moveTo(x=5723, y=-451) #printer icon
    time.sleep(0.1)
    pyautogui.click(clicks=2) 
    time.sleep(0.8)

    pyautogui.press(["left", "left"]) #navigate to left
    pyautogui.press(["up", "up", "right"]) #navigate to top row
    time.sleep(0.2)
    pyautogui.press("enter") #scroll left
    time.sleep(0.7)

    pyautogui.click(x=1619, y=542) #click on excel sheet
    time.sleep(0.5)
    pyautogui.press("down") #scroll down (press down key)
    pyautogui.hotkey('ctrl', 'c') #copy (ctrl c)
    time.sleep(0.2)
    pyautogui.moveTo(x=3456, y=42)
    time.sleep(0.15)
    pyautogui.click() #put cursor in file explorer
    time.sleep(0.2)
    pyautogui.hotkey('ctrl', 'v') #paste it (ctrl v)
    pyautogui.press("enter") #save it
    time.sleep(0.4)
    pyautogui.moveTo(x=5736, y=-490) #close trial balance
    time.sleep(0.1)
    pyautogui.click()
    time.sleep(0.5)