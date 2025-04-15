import cv2
import numpy as np
import mss
import pyautogui
pyautogui.FAILSAFE = True  # Lascia attivo il fail-safe, ma lo gestiamo sotto
import logging
import os
import sys
import time
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

# Log con nome univoco basato su timestamp
log_filename = f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
log_path = os.path.join(output_dir, log_filename)
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler(log_path, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# Immagine dei bottoni da cercare
template_radio = os.path.join(resource_path(), "radio.png")
template_buttons = [
    os.path.join(resource_path(), "button-grey.png"),
    os.path.join(resource_path(), "button.png")
]

# Verifica che i file template esistano
all_templates = [template_radio] + template_buttons
missing_files = [f for f in all_templates if not os.path.isfile(f)]

if missing_files:
    error_message = "❌ File template mancanti:\n" + "\n".join(f"- {f}" for f in missing_files)
    print(error_message)
    logging.critical(error_message)
    import ctypes
    ctypes.windll.user32.MessageBoxW(0, error_message, "Errore critico", 0x10)  # 0x10 = MB_ICONERROR
    sys.exit(1)
    
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
    try:
        x, y = pyautogui.position()
        pyautogui.moveTo(x + 1, y)
        pyautogui.moveTo(x - 1, y)
        pyautogui.moveTo(x, y)
        pyautogui.press('shift')
    except Exception as e:
        logging.warning(f"⚠️ Errore in avoid_standby: {e}")

# Screenshot della finestra
def capture_window_screenshot(hwnd, title, prefix=""):
    try:
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
            logging.info(f"✅ Screenshot salvato: {filename}")

        return img_bgr, (x, y)
    except Exception as e:
        logging.error(f"❌ Errore durante lo screenshot della finestra '{title}': {e}")
        return None, None

# Trova un template nell'immagine (se più match, clicca quello più in alto)
def find_template_position(img_bgr_original, template_path, threshold=THRESHOLD, scales=SCALES):
    try:
        template = cv2.imread(template_path, cv2.IMREAD_COLOR)
        if template is None:
            logging.error(f"❌ Template '{template_path}' non trovato.")
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
            logging.info(f"🎯 Match più in alto a y={best_y}, scala={best_scale:.2f}")
            return best_match, (scaled_w, scaled_h)
        else:
            logging.warning(f"❌ Nessun match sopra soglia {threshold} per {template_path}")
            return None
    except Exception as e:
        logging.error(f"❌ Errore durante il matching del template '{template_path}': {e}")
        return None

# Click relativo alla finestra
def click_at_position(screen_origin, template_pos, template_size):
    try:
        x0, y0 = screen_origin
        tx, ty = template_pos
        tw, th = template_size
        click_x = round(x0 + tx + tw / 2)
        click_y = round(y0 + ty + th / 2)
        logging.info(f"🖱️ Clic in ({click_x}, {click_y})")
        pyautogui.click(click_x, click_y)
    except Exception as e:
        logging.error(f"❌ Errore durante il clic: {e}")

# Processa una finestra: screenshot + interazioni
def process_window_interaction(hwnd, title):
    # Ripristina finestra se minimizzata; metti in primo piano; massimizza
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.2)
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            logging.warning(f"⚠️ SetForegroundWindow fallito: {e}")
            try:
                import ctypes
                logging.info("↪️ Provo fallback con SwitchToThisWindow...")
                ctypes.windll.user32.SwitchToThisWindow(hwnd, True)
            except Exception as e2:
                logging.error(f"❌ Fallback con SwitchToThisWindow fallito: {e2}")
                return
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        time.sleep(0.2)
    except Exception as e:
        logging.warning(f"⚠️ Errore nel gestire la finestra '{title}': {e}")
        return
    
    img_bgr, origin = capture_window_screenshot(hwnd, title)

    # Trova e clicca radio button
    try:
        radio = find_template_position(img_bgr, template_radio)

        if radio:
            logging.info("✅ Radio button 1 trovato.")
            click_at_position(origin, *radio)
        else:
            logging.warning("⚠️ Nessun radio button trovato.")
            return
    except Exception as e:
        logging.error(f"❌ Errore durante ricerca o clic radio: {e}")
        return

    time.sleep(1)

    # Trova e clicca il bottone "Invia". Cerca su diversi possibili template.
    # Nota: è necessario cercare il tasto ingrigito (diventa blu solo dopo che è stato selezionato il radio)
    #       perché al momento dello screenshot della finestra il tasto era ancora grigio. Tuttavia la posizione non cambia.
    try:
        for template in template_buttons:
            button = find_template_position(img_bgr, template)
            if button:
                break

        if button:
            logging.info("✅ Bottone trovato.")
            wait_time = np.random.uniform(0, 20)
            logging.info(f"⏳ Attendo {wait_time:.2f} secondi prima del clic...")
            time.sleep(wait_time)
            capture_window_screenshot(hwnd, title, prefix="CLICKED_")
            click_at_position(origin, *button)
        else:
            logging.warning("🚫 Bottone invia non trovato.")
    except Exception as e:
        logging.error(f"❌ Errore durante il clic del bottone: {e}")


# Loop principale
logging.info(f"🕵 Monitoraggio finestre per: '{keywords}'")
try:
    while True:
        matching_windows = [w for w in gw.getAllWindows()
                            if any(kw in w.title.lower() for kw in keywords)]
        if matching_windows:
            play_trill()
            for w in matching_windows:
                try:
                    logging.info(f"▶️ Finestra trovata: '{w.title}'")
                    process_window_interaction(w._hWnd, w.title)
                except pyautogui.FailSafeException:
                    logging.warning("🚨 Fail-safe attivato (mouse all'angolo).")
                    continue
                except Exception as e:
                    logging.error(f"⚠️ Errore durante l'interazione con la finestra '{w.title}': {e}")
        else:
            logging.info("❌ Nessuna finestra trovata.")
        avoid_standby()
        time.sleep(check_interval)
except KeyboardInterrupt:
    logging.info("🛑 Monitoraggio interrotto manualmente.")
    log_file.close()

