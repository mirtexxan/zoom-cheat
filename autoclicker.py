import cv2
import numpy as np
import mss
import os
import time
import pyautogui
import win32gui
import win32con
import pygetwindow as gw
import winsound
from datetime import datetime
import win32process
import win32api

# Finestra da cercare
search_string = "Sondaggi or Poll"
keywords = [k.strip().lower() for k in search_string.split("or")]

# Tempo tra un controllo e l'altro
check_interval = 5
n_beeps = 5

# Cartella contenente gli screenshot
output_dir = "screenshots"
os.makedirs(output_dir, exist_ok=True)

# Immagine del bottone da cercare
template_button="button.png"
template_radio = "radio.png"


# Trillo sonoro
def play_trill():
    for _ in range(n_beeps):
        winsound.Beep(3000, 200)
        time.sleep(0.1)

def sanitize_filename(title):
    return "".join(c if c.isalnum() else "_" for c in title)[:50]

# 📸 Screenshot della finestra
def capture_window_screenshot(hwnd, title):
    # Ripristina finestra se minimizzata
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    time.sleep(0.2)

    # FORZA IL FOCUS sulla finestra target
    try:
        fg_window = win32gui.GetForegroundWindow()
        current_thread_id = win32api.GetCurrentThreadId()
        fg_thread_id = win32process.GetWindowThreadProcessId(fg_window)[0]
        target_thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]

        win32gui.AttachThreadInput(current_thread_id, fg_thread_id, True)
        win32gui.AttachThreadInput(current_thread_id, target_thread_id, True)

        win32gui.SetForegroundWindow(hwnd)
        win32gui.SetFocus(hwnd)

        win32gui.AttachThreadInput(current_thread_id, fg_thread_id, False)
        win32gui.AttachThreadInput(current_thread_id, target_thread_id, False)
    except Exception as e:
        print(f"⚠️ Impossibile forzare il focus: {e}")

    # Massimizza dopo il focus
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    time.sleep(0.5)

    x, y, r, b = win32gui.GetWindowRect(hwnd)
    width, height = r - x, b - y

    with mss.mss() as sct:
        monitor = {"left": x, "top": y, "width": width, "height": height}
        sct_img = sct.grab(monitor)
        img = np.array(sct_img)
        img_bgr = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = sanitize_filename(title)
        filename = os.path.join(output_dir, f"{safe_title}_{timestamp}.png")
        cv2.imwrite(filename, img_bgr)
        print(f"✅ Screenshot salvato: {filename}")

    return img_bgr, (x, y)

# Trova un template nell'immagine
def find_template_position(img_bgr, template_path, threshold=0.8, scales=np.linspace(0.5, 2, 15)):
    template_orig = cv2.imread(template_path, cv2.IMREAD_COLOR)
    if template_orig is None:
        print(f"❌ Template '{template_path}' non trovato.")
        return None

    best_val = -1
    best_loc = None
    best_size = None

    for scale in scales:
        template = cv2.resize(template_orig, (0, 0), fx=scale, fy=scale, interpolation=cv2.INTER_LINEAR)
        if template.shape[0] > img_bgr.shape[0] or template.shape[1] > img_bgr.shape[1]:
            continue  # salta se il template è più grande dell'immagine

        result = cv2.matchTemplate(img_bgr, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)

        if max_val > best_val:
            best_val = max_val
            best_loc = max_loc
            best_size = template.shape[1], template.shape[0]  # width, height
            best_scale = scale

    if best_val >= threshold:
        print(f"🎯 Trovato con scala {best_scale:.2f}, confidenza: {best_val:.2f}")
        return best_loc, best_size
    else:
        print(f"❌ Nessuna corrispondenza sopra soglia ({threshold}) trovata.")
        return None

# 🖱️ Click relativo alla finestra
def click_at_position(screen_origin, template_pos, template_size):
    x0, y0 = screen_origin
    tx, ty = template_pos
    tw, th = template_size
    click_x = x0 + tx + tw // 2
    click_y = y0 + ty + th // 2
    print(f"🖱️ Clic in ({click_x}, {click_y})")
    pyautogui.click(click_x, click_y)

# 🤖 Processa una finestra: screenshot + interazioni
def process_window_interaction(hwnd, title):
    img_bgr, origin = capture_window_screenshot(hwnd, title)

    # Trova e clicca radio button (priorità a radio1)
    radio = find_template_position(img_bgr, template_radio)

    if radio:
        print("✅ Radio button 1 trovato.")
        click_at_position(origin, *radio)
        selected_radio = "radio"
    else:
        print("⚠️ Nessun radio button trovato.")
        return

    time.sleep(0.3)

    # Trova e clicca bottone solo se un radio è stato selezionato
    button = find_template_position(img_bgr, template_button)
    if button and selected_radio:
        print("✅ Bottone trovato. Procedo con click.")
        click_at_position(origin, *button)
    else:
        print("🚫 Bottone non cliccato: serve radio selezionato.")

# 🔁 Loop principale
print(f"🕵️ Monitoraggio finestre per: '{search_string}'")
try:
    while True:
        matching_windows = [w for w in gw.getAllWindows()
                            if any(kw in w.title.lower() for kw in keywords)]
        if matching_windows:
            play_trill()
            for w in matching_windows:
                print(f"\n▶️ Finestra trovata: '{w.title}'")
                process_window_interaction(w._hWnd, w.title)
        else:
            print("❌ Nessuna finestra trovata.")
        time.sleep(check_interval)
except KeyboardInterrupt:
    print("\n🛑 Monitoraggio interrotto.")
