import time
import pyautogui as gui


time.sleep(2)

gui.press('winleft')
gui.write("notepad", interval=0.03)
gui.press('enter')
time.sleep(1.5)  

gui.write("Olá! Este texto foi gerado por automação.", interval=0.05)
gui.press('enter')
gui.write("Ação realizada com pyautogui baseada no exemplo do e-mail.")
time.sleep(1)

gui.hotkey('ctrl', 's')
time.sleep(1)

gui.write("automacao_pyautogui.txt", interval=0.03)
gui.press('enter')
time.sleep(1)

gui.hotkey('alt', 'f4')
print("Automação concluída!")