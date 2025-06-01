import time
import tkinter as tk
from tkinter import ttk, messagebox
from pywinauto import Application
from pywinauto.keyboard import send_keys
from pywinauto.application import WindowSpecification
from pywinauto import Desktop
import threading
import win32con
import win32gui
import cv2
import numpy as np
from PIL import Image, ImageGrab
import os
import tempfile


class CoinPokerController:
    def __init__(self, root):
        self.root = root
        self.root.title("CoinPoker Controller")
        self.root.geometry("550x900")  # Larghezza aumentata a 500px
        self.root.resizable(False, False)
        
        # Variabili di stato
        self.is_running = False
        self.is_automating = False
        self.monitor_thread = None
        self.main_thread = None
        self.lobby_window = None
        self.auto_move_world = True
        self.world_count = 0
        self.ocr_enabled = True
        self.ocr_threshold = 0.6
        self.iteration_counter = 0  # Contatore iterazioni
        
        # Inizializza variabile per immagine target
        self.target_image = None
        
        # Configurazione GUI
        self.setup_gui()
        
        # Carica l'immagine di riferimento per OCR
        self.load_target_image()
        
        # Avvia il monitoraggio delle finestre World
        self.start_world_monitoring()
    
    def load_target_image(self):
        """Carica l'immagine Tavolo_chiuso.png"""
        try:
            if os.path.exists("Tavolo_chiuso.png"):
                self.target_image = cv2.imread("Tavolo_chiuso.png", cv2.IMREAD_COLOR)
                self.log_message("Immagine Tavolo_chiuso.png caricata con successo")
                self.log_message(f"Soglia OCR impostata a: {self.ocr_threshold}")
            else:
                self.log_message("ATTENZIONE: File Tavolo_chiuso.png non trovato nella directory corrente")
                messagebox.showwarning("Immagine mancante", 
                                     "File Tavolo_chiuso.png non trovato.\nPosizionare il file nella stessa directory del programma.")
        except Exception as e:
            self.log_message(f"Errore nel caricamento dell'immagine: {e}")
    
    def setup_gui(self):
        # Frame principale con padding aumentato
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Titolo
        title_label = ttk.Label(main_frame, text="CoinPoker Controller", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Status, contatori e iterazioni
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=1, column=0, columnspan=2, pady=(0, 20))
        
        self.status_label = ttk.Label(status_frame, text="Status: Fermo", 
                                     font=("Arial", 10))
        self.status_label.grid(row=0, column=0, columnspan=2, pady=(0, 5))
        
        self.world_count_label = ttk.Label(status_frame, text="Finestre World: 0", 
                                          font=("Arial", 10, "bold"), foreground="blue")
        self.world_count_label.grid(row=1, column=0, columnspan=2, pady=(0, 5))
        
        # Contatore iterazioni
        self.iteration_label = ttk.Label(status_frame, text="Iterazioni: 0", 
                                        font=("Arial", 12, "bold"), foreground="red")
        self.iteration_label.grid(row=2, column=0, columnspan=2)
        
        # Bottoni principali
        self.start_button = ttk.Button(main_frame, text="Avvia", 
                                      command=self.start_automation, width=20)
        self.start_button.grid(row=2, column=0, padx=(0, 10), pady=5)
        
        self.stop_button = ttk.Button(main_frame, text="Ferma", 
                                     command=self.stop_automation, 
                                     state="disabled", width=20)
        self.stop_button.grid(row=2, column=1, pady=5)
        
        # Bottoni Lobby
        lobby_frame = ttk.LabelFrame(main_frame, text="Controllo Lobby", padding="10")
        lobby_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        self.lobby_in_button = ttk.Button(lobby_frame, text="Lobby In (0,0)", 
                                         command=self.bring_lobby_in, width=22)
        self.lobby_in_button.grid(row=0, column=0, padx=(0, 10), pady=5)
        
        self.lobby_out_button = ttk.Button(lobby_frame, text="Lobby Out (-2000,0)", 
                                          command=self.move_lobby_out, width=22)
        self.lobby_out_button.grid(row=0, column=1, pady=5)
        
        # Controlli World Windows
        world_frame = ttk.LabelFrame(main_frame, text="Controllo Tavoli World", padding="10")
        world_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        self.auto_move_button = ttk.Button(world_frame, text="Auto-Sposta: ON", 
                                          command=self.toggle_auto_move, width=22)
        self.auto_move_button.grid(row=0, column=0, padx=(0, 10), pady=5)
        
        self.move_tables_button = ttk.Button(world_frame, text="Tavoli a (600,0)", 
                                           command=self.move_all_tables_to_position, width=22)
        self.move_tables_button.grid(row=0, column=1, pady=5)
        
        # Controlli OCR
        ocr_frame = ttk.LabelFrame(main_frame, text="Controllo OCR", padding="10")
        ocr_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        self.ocr_button = ttk.Button(ocr_frame, text="OCR: ON", 
                                    command=self.toggle_ocr, width=22)
        self.ocr_button.grid(row=0, column=0, padx=(0, 10), pady=5)
        
        self.test_ocr_button = ttk.Button(ocr_frame, text="Test OCR Tavoli", 
                                         command=self.test_ocr_manual, width=22)
        self.test_ocr_button.grid(row=0, column=1, pady=5)
        
        # Area di log
        log_label = ttk.Label(main_frame, text="Log:")
        log_label.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(20, 5))
        
        # Frame per il log con scrollbar
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.log_text = tk.Text(log_frame, height=22, width=60, wrap=tk.WORD)  # Dimensioni aumentate
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurazione del grid
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
    def update_iteration_counter(self):
        """Aggiorna il contatore delle iterazioni"""
        self.iteration_counter += 1
        self.iteration_label.config(text=f"Iterazioni: {self.iteration_counter}")
        
    def log_message(self, message):
        """Aggiunge un messaggio al log"""
        try:
            timestamp = time.strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            self.log_text.see(tk.END)
            self.root.update_idletasks()
        except AttributeError:
            print(f"[{timestamp}] {message}")
        
    def update_world_count(self, count):
        """Aggiorna il contatore delle finestre World"""
        self.world_count = count
        self.world_count_label.config(text=f"Finestre World: {count}")
        
    def toggle_auto_move(self):
        """Alterna l'interruttore per lo spostamento automatico"""
        self.auto_move_world = not self.auto_move_world
        status = "ON" if self.auto_move_world else "OFF"
        self.auto_move_button.config(text=f"Auto-Sposta: {status}")
        self.log_message(f"Spostamento automatico tavoli: {status}")
        
    def toggle_ocr(self):
        """Alterna l'interruttore per il controllo OCR"""
        self.ocr_enabled = not self.ocr_enabled
        status = "ON" if self.ocr_enabled else "OFF"
        self.ocr_button.config(text=f"OCR: {status}")
        self.log_message(f"Controllo OCR: {status}")
        
    def is_window_valid(self, window):
        """Verifica se la finestra è ancora valida"""
        try:
            window.window_text()
            window.rectangle()
            return True
        except Exception:
            return False
    
    def capture_window_screenshot_improved(self, window):
        """Cattura screenshot senza spostare il mouse"""
        screenshot_path = None
        try:
            if not self.is_window_valid(window):
                self.log_message("Finestra non valida per screenshot")
                return None, None
            
            # NON utilizzare set_focus() che può spostare il mouse
            # Usa solo restore() per de-minimizzare se necessario
            try:
                window.restore()
            except:
                pass  # Ignora errori se la finestra è già visibile
                
            time.sleep(1.5)
            
            # Ottieni le coordinate della finestra
            rect = window.rectangle()
            self.log_message(f"Screenshot finestra: {rect.left},{rect.top} - {rect.right},{rect.bottom}")
            
            # Cattura lo screenshot dell'area della finestra
            screenshot = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
            
            temp_dir = tempfile.gettempdir()
            screenshot_path = os.path.join(temp_dir, f"coinpoker_screenshot_{int(time.time())}.png")
            screenshot.save(screenshot_path)
            self.log_message(f"Screenshot salvato: {screenshot_path}")
            
            screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            return screenshot_cv, screenshot_path
            
        except Exception as e:
            self.log_message(f"Errore nella cattura screenshot: {e}")
            return None, screenshot_path
    
    def find_image_in_window_improved(self, window, window_title=""):
        """Cerca l'immagine target nella finestra del tavolo"""
        if self.target_image is None:
            self.log_message("Immagine target non caricata")
            return False
            
        screenshot_path = None
        try:
            window_screenshot, screenshot_path = self.capture_window_screenshot_improved(window)
            if window_screenshot is None:
                self.log_message(f"Impossibile catturare screenshot per: {window_title}")
                return False
            
            h_target, w_target = self.target_image.shape[:2]
            h_screen, w_screen = window_screenshot.shape[:2]
            self.log_message(f"Target: {w_target}x{h_target}, Screenshot: {w_screen}x{h_screen}")
            
            result = cv2.matchTemplate(window_screenshot, self.target_image, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            self.log_message(f"OCR Match - Min: {min_val:.3f}, Max: {max_val:.3f}, Soglia: {self.ocr_threshold}")
            
            if max_val >= self.ocr_threshold:
                self.log_message(f"🎯 IMMAGINE TROVATA! Confidenza: {max_val:.3f}")
                return True
            else:
                self.log_message(f"❌ Immagine non trovata. Confidenza: {max_val:.3f}")
                return False
                
        except Exception as e:
            self.log_message(f"Errore nell'analisi OCR: {e}")
            return False
        finally:
            if screenshot_path and os.path.exists(screenshot_path):
                threading.Timer(30.0, lambda: self.cleanup_screenshot(screenshot_path)).start()
    
    def cleanup_screenshot(self, screenshot_path):
        """Pulisce il file screenshot temporaneo"""
        try:
            if os.path.exists(screenshot_path):
                os.remove(screenshot_path)
                self.log_message(f"Screenshot temporaneo rimosso: {os.path.basename(screenshot_path)}")
        except Exception as e:
            self.log_message(f"Errore nella pulizia screenshot: {e}")
    
    def close_table_window_improved(self, window, window_title=""):
        """Chiude il tavolo senza spostare il mouse"""
        try:
            if not self.is_window_valid(window):
                self.log_message(f"❌ Finestra non più valida: {window_title}")
                return False
            
            self.log_message(f"Tentativo chiusura tavolo: {window_title}")
            
            # NON usare set_focus() che può spostare il mouse
            # Prova direttamente con i messaggi di sistema
            success = False
            
            # Metodo 1: WM_CLOSE diretto (più affidabile e senza mouse)
            try:
                hwnd = window.handle
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                time.sleep(2)
                if not self.is_window_valid(window):
                    self.log_message(f"✓ Tavolo chiuso con WM_CLOSE: {window_title}")
                    success = True
            except Exception as e:
                self.log_message(f"Errore WM_CLOSE: {e}")
            
            # Metodo 2: Alt+F4 solo se WM_CLOSE fallisce
            if not success:
                try:
                    # Porta brevemente il focus solo per inviare il comando
                    hwnd = window.handle
                    win32gui.PostMessage(hwnd, win32con.WM_SYSKEYDOWN, win32con.VK_F4, 0x20000000)
                    time.sleep(0.1)
                    win32gui.PostMessage(hwnd, win32con.WM_SYSKEYUP, win32con.VK_F4, 0x20000000)
                    time.sleep(2)
                    if not self.is_window_valid(window):
                        self.log_message(f"✓ Tavolo chiuso con Alt+F4: {window_title}")
                        success = True
                except Exception as e:
                    self.log_message(f"Errore Alt+F4: {e}")
            
            return success
            
        except Exception as e:
            self.log_message(f"Errore nella chiusura del tavolo: {e}")
            return False
    
    def scan_and_close_tables(self):
        """Scansiona tutti i tavoli e chiude quelli che contengono l'immagine target"""
        if not self.ocr_enabled:
            self.log_message("OCR disabilitato - skip controllo tavoli")
            return
            
        try:
            world_windows = self.find_world_windows_alternative()
            if not world_windows:
                self.log_message("Nessun tavolo World trovato per il controllo OCR")
                return
            
            self.log_message(f"🔍 Inizio controllo OCR su {len(world_windows)} tavoli")
            
            tables_to_close = []
            
            for i, window in enumerate(world_windows):
                if not self.is_automating:
                    break
                
                try:
                    window_title = window.window_text()
                    self.log_message(f"--- Controllo tavolo {i+1}/{len(world_windows)}: {window_title} ---")
                    
                    if self.find_image_in_window_improved(window, window_title):
                        self.log_message(f"⚠️ TROVATA immagine target nel tavolo: {window_title}")
                        tables_to_close.append((window, window_title))
                    else:
                        self.log_message(f"✅ Tavolo OK: {window_title}")
                    
                    if i < len(world_windows) - 1:
                        time.sleep(3)
                        
                except Exception as e:
                    self.log_message(f"Errore nel controllo tavolo {i+1}: {e}")
            
            if tables_to_close:
                self.log_message(f"🚫 Trovati {len(tables_to_close)} tavoli da chiudere")
                closed_count = 0
                
                for window, window_title in tables_to_close:
                    if not self.is_automating:
                        break
                    
                    self.log_message(f"Chiusura tavolo: {window_title}")
                    if self.close_table_window_improved(window, window_title):
                        closed_count += 1
                    
                    time.sleep(3)
                
                self.log_message(f"✅ Tavoli chiusi: {closed_count}/{len(tables_to_close)}")
            else:
                self.log_message("✅ Nessun tavolo da chiudere")
                
        except Exception as e:
            self.log_message(f"Errore nel controllo OCR: {e}")
    
    def test_ocr_manual(self):
        """Test manuale del controllo OCR"""
        self.log_message("🧪 === INIZIO TEST OCR MANUALE ===")
        threading.Thread(target=self.scan_and_close_tables, daemon=True).start()
        
    def move_all_tables_to_position(self):
        """Sposta tutti i tavoli World alle coordinate 600,0"""
        try:
            world_windows = self.find_world_windows_alternative()
            if world_windows:
                self.log_message(f"Spostamento di {len(world_windows)} tavoli a (600,0)")
                for window in world_windows:
                    try:
                        window.move_window(600, 0)
                        self.log_message(f"Tavolo spostato: {window.window_text()}")
                    except Exception as e:
                        self.log_message(f"Errore nello spostamento tavolo: {e}")
            else:
                self.log_message("Nessun tavolo World trovato")
        except Exception as e:
            self.log_message(f"Errore nel trovare i tavoli: {e}")
        
    def find_coinpoker_window(self):
        """Trova la finestra CoinPoker"""
        try:
            app = Application(backend="win32").connect(title_re="CoinPoker - Lobby", 
                                                     class_name_re="Qt673QWindowIcon")
            lobby_window = app.window(title_re="CoinPoker - Lobby")
            return lobby_window
        except Exception as e:
            self.log_message(f"Errore nel trovare la finestra CoinPoker: {e}")
            return None
    
    def find_world_windows_alternative(self):
        """Trova tutte le finestre con 'World' nel titolo"""
        world_windows = []
        
        def enum_windows_callback(hwnd, windows):
            try:
                window_title = win32gui.GetWindowText(hwnd)
                if "World" in window_title and win32gui.IsWindowVisible(hwnd):
                    app = Application(backend="win32").connect(handle=hwnd)
                    window = app.window(handle=hwnd)
                    if self.is_window_valid(window):
                        windows.append(window)
            except Exception:
                pass
            return True
        
        try:
            win32gui.EnumWindows(enum_windows_callback, world_windows)
        except Exception as e:
            self.log_message(f"Errore nell'enumerazione delle finestre: {e}")
        
        return world_windows
    
    def send_key_to_window_background(self, window, key_code):
        """Invia un tasto senza spostare il mouse o il focus"""
        try:
            hwnd = window.handle
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, key_code, 0)
            time.sleep(0.05)
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, key_code, 0)
            return True
        except Exception as e:
            self.log_message(f"Errore nell'invio del tasto: {e}")
            return False
    
    def press_enter(self, lobby_window):
        """Simula Enter senza spostare il mouse"""
        # Usa direttamente PostMessage per evitare movimenti del mouse
        try:
            hwnd = lobby_window.handle
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            time.sleep(0.05)
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        except Exception as e:
            self.log_message(f"Errore Enter: {e}")
    
    def press_down(self, lobby_window):
        """Simula freccia giù senza spostare il mouse"""
        try:
            hwnd = lobby_window.handle
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
            time.sleep(0.05)
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        except Exception as e:
            self.log_message(f"Errore Down: {e}")
    
    def press_page_up(self):
        """Simula PageUp usando SendInput invece di pywinauto"""
        try:
            # Usa PostMessage diretto se possibile, altrimenti send_keys
            send_keys('{PGUP}')
            time.sleep(1)
        except Exception as e:
            self.log_message(f"Errore PageUp: {e}")
    
    def update_status(self, status):
        """Aggiorna lo status"""
        self.status_label.config(text=f"Status: {status}")
        
    def move_window_offscreen(self, window):
        """Sposta la finestra fuori schermo se auto-move è attivo"""
        if not self.auto_move_world:
            return
            
        try:
            window.move_window(-1000, -1000)
            self.log_message(f"Finestra spostata fuori schermo: {window.window_text()}")
        except Exception as e:
            self.log_message(f"Errore nello spostamento della finestra: {e}")
    
    def monitor_world_windows(self):
        """Monitora e sposta le finestre World"""
        self.is_running = True
        while self.is_running:
            try:
                world_windows = self.find_world_windows_alternative()
                count = len(world_windows)
                
                self.update_world_count(count)
                
                if world_windows and self.auto_move_world:
                    for window in world_windows:
                        self.move_window_offscreen(window)
                        
            except Exception as e:
                self.log_message(f"Errore nel monitoraggio delle finestre: {e}")
            time.sleep(3)
    
    def start_world_monitoring(self):
        """Avvia il monitoraggio delle finestre World"""
        if not hasattr(self, 'monitor_thread') or not self.monitor_thread or not self.monitor_thread.is_alive():
            self.monitor_thread = threading.Thread(target=self.monitor_world_windows, daemon=True)
            self.monitor_thread.start()
            self.log_message("Monitoraggio finestre World avviato")
    
    def automation_cycle(self):
        """Ciclo principale di automazione con contatore"""
        while self.is_automating:
            try:
                if not self.lobby_window:
                    self.lobby_window = self.find_coinpoker_window()
                    if not self.lobby_window:
                        self.log_message("Impossibile trovare la finestra CoinPoker")
                        time.sleep(5)
                        continue
                
                self.log_message("Inizio ciclo di automazione")
                
                # 50 iterazioni di Enter + Down con contatore in tempo reale
                for i in range(50):
                    if not self.is_automating:
                        break
                    
                    # Aggiorna il contatore in tempo reale
                    self.update_iteration_counter()
                    
                    if i % 10 == 0:
                        self.log_message(f"Iterazione {i+1}/100 (Totale: {self.iteration_counter})")
                    
                    self.press_enter(self.lobby_window)
                    time.sleep(1)
                    self.press_down(self.lobby_window)
                    time.sleep(1)
                
                if not self.is_automating:
                    break
                
                self.log_message("🔍 === INIZIO CONTROLLO OCR TAVOLI ===")
                self.scan_and_close_tables()
                self.log_message("✅ === FINE CONTROLLO OCR TAVOLI ===")
                
                if not self.is_automating:
                    break
                
                # PageUp senza spostare il focus in modo aggressivo
                self.log_message("Invio comandi PageUp")
                for i in range(10):
                    if not self.is_automating:
                        break
                    self.log_message(f"PageUp {i+1}/10")
                    self.press_page_up()
                
                self.log_message("Ciclo completato")
                time.sleep(5)
                
            except Exception as e:
                self.log_message(f"Errore nel ciclo principale: {e}")
                time.sleep(5)
    
    def start_automation(self):
        """Avvia l'automazione"""
        if self.is_automating:
            return
            
        self.is_automating = True
        self.iteration_counter = 0  # Reset contatore
        self.update_iteration_counter()
        self.log_message("Avvio automazione...")
        self.update_status("In esecuzione")
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        
        self.main_thread = threading.Thread(target=self.automation_cycle, daemon=True)
        self.main_thread.start()
    
    def stop_automation(self):
        """Ferma l'automazione"""
        self.log_message("Fermando automazione...")
        self.is_automating = False
        self.update_status("Fermo")
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
    
    def bring_lobby_in(self):
        """Posiziona la lobby alle coordinate 0,0"""
        try:
            if not self.lobby_window:
                self.lobby_window = self.find_coinpoker_window()
            
            if self.lobby_window:
                self.lobby_window.move_window(0, 0)
                self.lobby_window.restore()
                self.log_message("Lobby posizionata alle coordinate 0,0")
            else:
                self.log_message("Impossibile trovare la finestra della lobby")
                messagebox.showerror("Errore", "Impossibile trovare la finestra CoinPoker")
        except Exception as e:
            self.log_message(f"Errore nel posizionare la lobby: {e}")

    
    def move_lobby_out(self):
        """Sposta la lobby alle coordinate -2000,0"""
        try:
            if not self.lobby_window:
                self.lobby_window = self.find_coinpoker_window()
            
            if self.lobby_window:
                self.lobby_window.move_window(-2000, 0)  # Sposta a -2000,0
                self.log_message("Lobby spostata alle coordinate -2000,0")
            else:
                self.log_message("Impossibile trovare la finestra della lobby")
                messagebox.showerror("Errore", "Impossibile trovare la finestra CoinPoker")
        except Exception as e:
            self.log_message(f"Errore nel spostare la lobby: {e}")
            messagebox.showerror("Errore", f"Errore: {e}")
    
    def on_closing(self):
        """Gestisce la chiusura dell'applicazione"""
        self.is_running = False
        self.is_automating = False
        time.sleep(1)  # Attendi che i thread si fermino
        self.root.destroy()


def main():
    root = tk.Tk()
    app = CoinPokerController(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()
