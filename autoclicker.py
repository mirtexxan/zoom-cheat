import cv2
import numpy as np
import mss
import os
import sys
import time
import pyautogui
import win32gui
import win32con
import pygetwindow as gw
import winsound
from datetime import datetime
import win32process
import win32api

# Parole chiave da cercare nel titolo della finestra
keywords = [
    "Sondaggi",
    "Poll",
]
keywords = [k.strip().lower() for k in keywords]

# Tempo tra un controllo e l'altro
check_interval = 5
n_beeps = 5

# Cartella contenente gli screenshot
def resource_path():
    if getattr(sys, 'frozen', False):
        # Se è un exe (pyinstaller)
        return os.path.dirname(sys.executable)
    else:
        # Se è uno script .py
        return os.path.dirname(os.path.abspath(__file__))

output_dir = os.path.join(resource_path(), "screenshots")
os.makedirs(output_dir, exist_ok=True)

# Immagine del bottone da cercare
template_radio = os.path.join(resource_path(), "radio.png")
template_button = os.path.join(resource_path(), "button.png")

# Parametri del template matching
SCALES = np.linspace(0.5, 2, 20)
THRESHOLD = 0.7

# Trillo sonoro
def play_trill():
    for _ in range(n_beeps):
        winsound.Beep(3000, 200)
        time.sleep(0.1)

def timestamp():
    return datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

def sanitize_filename(title):
    return "".join(c if c.isalnum() else "_" for c in title)[:50]

def avoid_standby():
    x, y = pyautogui.position()
    pyautogui.moveTo(x + 1, y)
    pyautogui.moveTo(x - 1, y)
    pyautogui.moveTo(x, y)
    pyautogui.press('shift')

# Screenshot della finestra
def capture_window_screenshot(hwnd, title, prefix=""):
    x, y, r, b = win32gui.GetWindowRect(hwnd)
    width, height = r - x, b - y

    with mss.mss() as sct:
        monitor = {"left": x, "top": y, "width": width, "height": height}
        sct_img = sct.grab(monitor)
        img = np.array(sct_img)
        img_bgr = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        
        safe_title = sanitize_filename(title)
        filename = f"{prefix}{timestamp()}_{safe_title}.png"
        filename = os.path.join(output_dir, filename)
        cv2.imwrite(filename, img_bgr)
        print(f"✅ Screenshot salvato: {filename}")

    return img_bgr, (x, y)

# Trova un template nell'immagine (se più match, clicca quello più in alto)
def find_template_position(img_bgr_original, template_path, threshold=THRESHOLD, scales=SCALES):
    template = cv2.imread(template_path, cv2.IMREAD_COLOR)
    if template is None:
        print(f"❌ Template '{template_path}' non trovato.")
        return None

    template_h, template_w = template.shape[:2]

    best_match = None
    best_y = float('inf')
    best_scale = 1.0

    for scale in scales:
        scaled_img = cv2.resize(img_bgr_original, (0, 0), fx=scale, fy=scale)
        if scaled_img.shape[0] < template_h or scaled_img.shape[1] < template_w:
            continue

        result = cv2.matchTemplate(scaled_img, template, cv2.TM_CCOEFF_NORMED)
        y_coords, x_coords = np.where(result >= threshold)

        for x, y in zip(x_coords, y_coords):
            corrected_x = int(x / scale)
            corrected_y = int(y / scale)

            if corrected_y < best_y:
                best_y = corrected_y
                best_match = (corrected_x, corrected_y)
                best_scale = scale

    if best_match:
        scaled_w = int(template_w / best_scale)
        scaled_h = int(template_h / best_scale)
        print(f"🎯 Match più in alto a y={best_y}, scala={best_scale:.2f}")
        return best_match, (scaled_w, scaled_h)
    else:
        print(f"❌ Nessun match sopra soglia {threshold} per {template_path}")
        return None

# Click relativo alla finestra
def click_at_position(screen_origin, template_pos, template_size):
    x0, y0 = screen_origin
    tx, ty = template_pos
    tw, th = template_size
    click_x = round(x0 + tx + tw / 2)
    click_y = round(y0 + ty + th / 2)
    print(f"🖱️ Clic in ({click_x}, {click_y})")
    pyautogui.click(click_x, click_y)

# Processa una finestra: screenshot + interazioni
def process_window_interaction(hwnd, title):
    # Ripristina finestra se minimizzata; metti in primo piano; massimizza
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    time.sleep(0.2)
    win32gui.SetForegroundWindow(hwnd)
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    time.sleep(0.2)
    
    img_bgr, origin = capture_window_screenshot(hwnd, title)

    # Trova e clicca radio button
    radio = find_template_position(img_bgr, template_radio)

    if radio:
        print("✅ Radio button 1 trovato.")
        click_at_position(origin, *radio)
    else:
        print("⚠️ Nessun radio button trovato.\n")
        return

    time.sleep(1)

    # Trova e clicca bottone
    button = find_template_position(img_bgr, template_button)
    if button:
        print("✅ Bottone trovato.")
        capture_window_screenshot(hwnd, title, prefix="CLICKED_")
        click_at_position(origin, *button)
        print("\n")
    else:
        print("🚫 Bottone invia non trovato.\n")


# Loop principale
print(f"🕵 Monitoraggio finestre per: '{keywords}'")
try:
    while True:
        matching_windows = [w for w in gw.getAllWindows()
                            if any(kw in w.title.lower() for kw in keywords)]
        if matching_windows:
            play_trill()
            for w in matching_windows:
                try:
                    print(f"\n[{timestamp()}] ▶️ Finestra trovata: '{w.title}'")
                    process_window_interaction(w._hWnd, w.title)
                except Exception as e:
                    print(f"⚠️ Errore durante l'interazione con la finestra '{w.title}': {e}\n")
        else:
            print(f"[{timestamp()}] ❌ Nessuna finestra trovata.")
        avoid_standby()
        time.sleep(check_interval)
except KeyboardInterrupt:
    print("\n🛑 Monitoraggio interrotto.")
