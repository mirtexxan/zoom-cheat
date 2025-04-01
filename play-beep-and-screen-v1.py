import mss
import cv2
import numpy as np
import win32gui
import win32con
import pygetwindow as gw
import time
import os
from datetime import datetime
import winsound
import win32process
import win32api
import win32com.client

# Titolo (o parte) della finestra da cercare
search_string = "Sondagg or Poll"
check_interval = 5  # secondi
keywords = [k.strip().lower() for k in search_string.split("or")]

# 📁 Cartella per salvare gli screenshot
output_dir = "."
os.makedirs(output_dir, exist_ok=True)

def play_trill():
    print("🔔 FINESTRA TROVATA - SUONO D'ALLARME!")
    for _ in range(5):  # Trillo ripetuto
        winsound.Beep(3000, 200)  # Frequenza 3000 Hz, durata 200 ms
        time.sleep(0.1)

def sanitize_filename(title):
    return "".join(c if c.isalnum() else "_" for c in title)[:50]

def capture_and_save_window(hwnd, title):
    # Ripristina se minimizzata
    if win32gui.IsIconic(hwnd):
        print(f"↩️ Ripristino finestra minimizzata: '{title}'")
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.2)

    # Massimizza
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    time.sleep(0.5)

    # Screenshot
    x, y, r, b = win32gui.GetWindowRect(hwnd)
    width, height = r - x, b - y

    with mss.mss() as sct:
        monitor = {"left": x, "top": y, "width": width, "height": height}
        try:
            sct_img = sct.grab(monitor)
            img = np.array(sct_img)
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_title = sanitize_filename(title)
            filename = os.path.join(output_dir, f"{safe_title}_{timestamp}.png")

            cv2.imwrite(filename, img)
            print(f"✅ Screenshot salvato: {filename}")
        except Exception as e:
            print(f"❌ Errore nel catturare '{title}': {e}")

# 🔁 Monitoraggio continuo
print(f"🕵️‍♂️ Monitoraggio continuo per: '{search_string}'")
try:
    while True:
        matching_windows = [w for w in gw.getAllWindows()
                            if any(kw in w.title.lower() for kw in keywords)]
        if matching_windows:
            for w in matching_windows:
                print(f"✅ Finestra trovata: '{w.title}'")
                capture_and_save_window(w._hWnd, w.title)
            play_trill()
        else:
            print("❌ Nessuna finestra trovata.")

        time.sleep(check_interval)

except KeyboardInterrupt:
    print("\n🛑 Monitoraggio interrotto.")
