import pyautogui
import time

print("Posicione o mouse no local desejado...")
time.sleep(3)  # Aguarda 3 segundos para você posicionar o mouse

x, y = pyautogui.position()
print(f"Coordenadas do mouse: X={x}, Y={y}")
