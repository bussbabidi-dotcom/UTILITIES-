import pyautogui
import time
time.sleep(3)
for i in range (1):
 x, y = pyautogui.position()
 print(f"Posição encontrada: X={x}, Y={y}")
 time.sleep(0.5)