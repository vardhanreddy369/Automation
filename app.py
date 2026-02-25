import customtkinter as ctk
import pyautogui
import time
import threading
import os
import platform
import subprocess
import json
from datetime import datetime
from tkinter import filedialog

# Securely import pywinauto ONLY if the OS is Windows
try:
    if platform.system() == "Windows":
        from pywinauto import Application
except ImportError:
    pass

# Set the appearance and color theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")  # Changed to 'green' for a different premium feel

class LegacyAutomationBot(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title("Practice Management Bot PRO")
        self.geometry(f"{600}x{650}")  # Increased height for settings
        self.resizable(False, False)

        # Settings Persistence
        self.config_file = "config.json"
        self.app_path = self.load_config()

        # ------------------ UI Layout ------------------
        # Header Label
        self.header_label = ctk.CTkLabel(self, text="🏥 Practice App Automation", font=ctk.CTkFont(size=24, weight="bold"))
        self.header_label.pack(pady=(20, 5))

        # Instructions
        self.instructions = ctk.CTkLabel(
            self, 
            text="Automates legacy medical/dental management software.\nChoose an automation engine below:", 
            text_color="gray"
        )
        self.instructions.pack(pady=(0, 20))

        # Input Frame for Patient Data
        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.pack(pady=10, padx=20, fill="x")
        
        self.fname_label = ctk.CTkLabel(self.input_frame, text="Patient First Name:")
        self.fname_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.fname_entry = ctk.CTkEntry(self.input_frame, width=150)
        self.fname_entry.grid(row=0, column=1, padx=10, pady=10)
        self.fname_entry.insert(0, "John")

        self.lname_label = ctk.CTkLabel(self.input_frame, text="Patient Last Name:")
        self.lname_label.grid(row=0, column=2, padx=10, pady=10, sticky="w")
        self.lname_entry = ctk.CTkEntry(self.input_frame, width=150)
        self.lname_entry.grid(row=0, column=3, padx=10, pady=10)
        self.lname_entry.insert(0, "Doe")

        # Status Display
        self.status_frame = ctk.CTkFrame(self, fg_color="black")
        self.status_frame.pack(pady=10, padx=20, fill="both", expand=True)

        self.status_label = ctk.CTkLabel(self.status_frame, text="Status: Ready", font=ctk.CTkFont(size=14), text_color="limegreen")
        self.status_label.pack(pady=20, padx=10)

        # --- SETTINGS SECTION ---
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.pack(pady=10, padx=20, fill="x")

        self.path_label = ctk.CTkLabel(self.settings_frame, text="Target Application Path:", font=ctk.CTkFont(weight="bold"))
        self.path_label.pack(pady=(10, 0), padx=10)

        self.path_display = ctk.CTkEntry(self.settings_frame, width=400)
        self.path_display.pack(pady=5, padx=10)
        self.path_display.insert(0, self.app_path)
        self.path_display.configure(state="disabled")

        self.browse_btn = ctk.CTkButton(
            self.settings_frame, 
            text="📁 Select App Executable", 
            command=self.browse_app_path,
            fg_color="#444", hover_color="#555"
        )
        self.browse_btn.pack(pady=(0, 10))

        # Action Buttons (PyWinAuto & PyAutoGUI)
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(pady=(10, 20), padx=20, fill="x")

        self.pywin_btn = ctk.CTkButton(
            self.button_frame, 
            text="⚡ Run PyWinAuto\n(Native Elements)", 
            font=ctk.CTkFont(size=14, weight="bold"),
            height=50,
            command=self.start_pywinauto_thread,
            fg_color="#005b96", hover_color="#03396c"
        )
        self.pywin_btn.pack(side="left", expand=True, fill="x", padx=5)

        self.pyauto_btn = ctk.CTkButton(
            self.button_frame, 
            text="📷 Run PyAutoGUI\n(Screen/Images)", 
            font=ctk.CTkFont(size=14, weight="bold"),
            height=50,
            command=self.start_pyautogui_thread,
            fg_color="#b33939", hover_color="#cd6133"
        )
        self.pyauto_btn.pack(side="right", expand=True, fill="x", padx=5)

    # ------------------ UI Helpers ------------------
    def update_status(self, text, color="limegreen"):
        self.after(0, lambda: self.status_label.configure(text=f"Status: {text}", text_color=color))

    def disable_buttons(self):
        self.after(0, lambda: self.pywin_btn.configure(state="disabled"))
        self.after(0, lambda: self.pyauto_btn.configure(state="disabled"))
        self.after(0, lambda: self.browse_btn.configure(state="disabled"))
        
    def enable_buttons(self):
        self.after(0, lambda: self.pywin_btn.configure(state="normal"))
        self.after(0, lambda: self.pyauto_btn.configure(state="normal"))
        self.after(0, lambda: self.browse_btn.configure(state="normal"))

    # ------------------ Config Persistence ------------------
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r") as f:
                    data = json.load(f)
                    return data.get("app_path", "winword")
            except:
                return "winword"
        return "winword"

    def save_config(self):
        with open(self.config_file, "w") as f:
            json.dump({"app_path": self.app_path}, f)

    def browse_app_path(self):
        file_path = filedialog.askopenfilename(
            title="Select Application Executable",
            filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
        )
        if file_path:
            self.app_path = file_path
            self.path_display.configure(state="normal")
            self.path_display.delete(0, "end")
            self.path_display.insert(0, self.app_path)
            self.path_display.configure(state="disabled")
            self.save_config()
            self.update_status(f"App Path Updated: {os.path.basename(file_path)}", color="cyan")

    # ------------------ PyWinAuto Logic (The Best Tool) ------------------
    def start_pywinauto_thread(self):
        fname = self.fname_entry.get()
        lname = self.lname_entry.get()
        self.disable_buttons()
        threading.Thread(target=self.run_pywinauto, args=(fname, lname), daemon=True).start()

    def run_pywinauto(self, fname, lname):
        if platform.system() != "Windows":
             self.update_status("Error: PyWinAuto ONLY works on Windows!", color="red")
             self.enable_buttons()
             return

        try:
            app_name = os.path.basename(self.app_path)
            self.update_status(f"Launching {app_name}...", color="yellow")
            
            # Start User Selected Process
            if self.app_path == "winword":
                subprocess.Popen(["cmd", "/c", "start winword"], shell=True)
            else:
                subprocess.Popen([self.app_path])

            self.update_status("Waiting for application to load...", color="yellow")
            time.sleep(7)
            
            # Connect to the process
            try:
                if self.app_path == "winword":
                    app = Application(backend="uia").connect(path="winword.exe", timeout=10)
                else:
                    app = Application(backend="uia").connect(path=self.app_path, timeout=10)
            except Exception:
                app = Application(backend="uia").connect(title_re=f".*{app_name.split('.')[0]}.*", timeout=10)

            # Bring whichever Word Window just launched to the Front
            self.update_status("Finding active Word window...", color="yellow")
            main_window = app.top_window()
            main_window.set_focus()
            
            # Step 2: Create the Blank Document
            self.update_status("Selecting 'Blank Document'...", color="yellow")
            pyautogui.press('enter')
            time.sleep(4)

            # Step 3: BRUTE FORCE FOCUS (Ensure we are typing in the new doc)
            self.update_status("Focusing newest document...", color="yellow")
            doc_window = app.top_window()
            doc_window.set_focus()
            time.sleep(1)
            
            self.update_status(f"Executing Visual Typing...", color="yellow")
            
            # Typing the patient data
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            typing_text = f"Automated Patient Entry (Pro Mode)\nFirst Name: {fname}\nLast Name: {lname}\nEntry Time: {timestamp}"
            pyautogui.write(typing_text, interval=0.05)
            time.sleep(1)

            # --- STEP 4: AUTOMATED SAVE SEQUENCE ---
            self.update_status("Saving Document...", color="cyan")
            # Open Save As dialog (F12 is the fastest way in Word)
            pyautogui.press('f12')
            time.sleep(2)
            
            # Type a unique filename based on the patient name
            safe_fname = "".join(x for x in fname if x.isalnum())
            safe_lname = "".join(x for x in lname if x.isalnum())
            filename = f"Patient_{safe_fname}_{safe_lname}_{timestamp}"
            
            pyautogui.write(filename)
            time.sleep(1)
            pyautogui.press('enter')
            
            self.update_status(f"Saved: {filename}.docx ✅", color="limegreen")
            time.sleep(2)
            
            self.update_status("Automation Cycle Complete! ✨", color="limegreen")
        except Exception as e:
            self.update_status(f"Execution Error: {e}", color="red")
        finally:
            self.enable_buttons()

    # ------------------ PyAutoGUI Logic (Screen Coordinate Fallback) ------------------
    def start_pyautogui_thread(self):
        self.disable_buttons()
        threading.Thread(target=self.run_pyautogui, daemon=True).start()

    def run_pyautogui(self):
        try:
            self.update_status("Launching App via Start Menu...", color="yellow")
            app_name = os.path.basename(self.app_path)
            
            if platform.system() == "Windows":
                 pyautogui.hotkey('win', 'r')
                 time.sleep(1)
                 # Use the user-defined path or winword
                 target = self.app_path if self.app_path != "winword" else "winword"
                 pyautogui.typewrite(target)
                 pyautogui.press('enter')
            elif platform.system() == "Darwin":
                 # For Mac, we use the selected path or default to Word
                 target_name = app_name if self.app_path != "winword" else "Microsoft Word"
                 self.update_status(f"Opening/Focusing {target_name}...", color="yellow")
                 os.system(f"open -a '{target_name}'")
                 time.sleep(1)
                 os.system(f"osascript -e 'tell application \"{target_name}\" to activate'")
                 time.sleep(2)
                 
            self.update_status(f"Waiting for {app_name} to load...", color="yellow")
            time.sleep(6) 
            
            self.update_status("Interacting with New Document...", color="yellow")
            if platform.system() == "Windows":
                 pyautogui.press('enter')
            elif platform.system() == "Darwin":
                 # Focus app one more time right before opening a document
                 target_name = app_name if self.app_path != "winword" else "Microsoft Word"
                 os.system(f"osascript -e 'tell application \"{target_name}\" to activate'")
                 pyautogui.hotkey('command', 'n')
            time.sleep(2)

            self.update_status("Clicking arbitrary start coordinates...", color="yellow")
            pyautogui.click(x=450, y=300)
            time.sleep(1)

            self.update_status("Searching for 'menu_button.png' on screen...", color="orange")
            try:
                # This will silently fail if 'menu_button.png' doesn't exist, so we catch it
                button = pyautogui.locateOnScreen('menu_button.png')
                if button:
                     pyautogui.click(button)
                     self.update_status("Clicked Image Match!", color="limegreen")
                else:
                     self.update_status("Image not found on screen.", color="orange")
            except Exception:
                self.update_status("Note: 'menu_button.png' missing from folder. Skipped image click.", color="orange")
            
            pyautogui.write("\nPyAutoGUI Fallback Sequence Finished.", interval=0.05)
            self.update_status("Screen Automation Complete! ✅", color="limegreen")

        except Exception as e:
            self.update_status(f"PyAutoGUI Error: {e}", color="red")
        finally:
            self.enable_buttons()

if __name__ == "__main__":
    pyautogui.FAILSAFE = True  # Move mouse to corner to abort
    app = LegacyAutomationBot()
    app.mainloop()
