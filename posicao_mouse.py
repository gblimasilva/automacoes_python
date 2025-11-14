import pyautogui
import time

print("🔹 Posicione o mouse sobre o botão da macro em 5 segundos...")
time.sleep(5)

posicao = pyautogui.position()
print(f"🖱️ Posição capturada: {posicao}")
