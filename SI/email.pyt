import getpass
import time
import pyautogui as gui


gui.FAILSAFE = True
gui.PAUSE = 0.5   

email = input("Digite seu email: ")
senha = getpass.getpass("Digite sua senha: ")

gui.press('winleft')
gui.write("edge", interval=0.02)
gui.press('enter')
time.sleep(0.5)
gui.hotkey("ctrl", "shift", "n")
time.sleep(0.5)

gui.write("gmail.com")
gui.press("enter")
time.sleep(3) 


gui.write(email)
gui.press("enter")
time.sleep(3)

gui.write(senha)
gui.press("enter")
time.sleep(5) 

gui.moveTo(135, 210)
time.sleep(1)
gui.click()
time.sleep(2)

gui.write("prof.humbertoltoledo@gmail.com")
gui.press("enter")
gui.press("tab")     
time.sleep(1)

# Escrever assunto
gui.write("Eai professor, conseguimos, Victor Marques")
time.sleep(1)
gui.press("tab")     

gui.hotkey("ctrl", "v")
time.sleep(2)         

gui.press("tab")    
gui.press("enter")
time.sleep(3)

gui.hotkey("alt", "f4")
time.sleep(1)
gui.hotkey("alt", "f4")