import pyautogui
from time import sleep


def soneca():
    return sleep(2)

pyautogui.press("win")
soneca()
pyautogui.write("Teams")
soneca()
pyautogui.press("enter")
sleep(30)
pyautogui.click(x=860, y=70)
sleep(8)
pyautogui.click(x=850, y=450)
soneca()
pyautogui.scroll(-450)
soneca()
pyautogui.rightClick(x=1000, y=598)
soneca()
pyautogui.click(x=1010, y=450)
soneca()
pyautogui.keyDown('alt')
soneca()
pyautogui.press("F4")
pyautogui.keyUp("alt")