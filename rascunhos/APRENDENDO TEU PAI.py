import pyautogui
import time

y = 245

for i in range (2):

    # Primeiro clique
    pyautogui.click(x=210, y=y, clicks=2, interval=0.1)

    # Seleciona e copia
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    # Segundo local
    pyautogui.click(x=1680, y=65, clicks=2, interval=0.1)

    # Cola
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'v')

    # Próximo ciclo: diminui 7 pixels no Y
    y -= 10

    # Espera 1 segundo antes de repetir
    time.sleep(1)