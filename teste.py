import pyautogui
import subprocess

pyautogui.press("win")
pyautogui.write("Teams")
subprocess.run( shell=True)
pyautogui.click(x=500, y=50)

pyautogui.press("enter")