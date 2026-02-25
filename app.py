import customtkinter as ctk
import pyautogui
import time
import threading
import os
import platform
import subprocess

# Securely import pywinauto ONLY if the OS is Windows
# (Otherwise, it will crash your Mac when trying to test the UI!)
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
        self.title("Practice Management Bot")
        self.geometry(f"{600}x{500}")
        self.resizable(False, False)

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
        
    def enable_buttons(self):
        self.after(0, lambda: self.pywin_btn.configure(state="normal"))
        self.after(0, lambda: self.pyauto_btn.configure(state="normal"))

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
            self.update_status("Starting PracticeApp (PyWinAuto)...", color="yellow")
            
            # THE EXACT CODE SUMAN PROVIDED
            # This points squarely at the legacy PracticeApp you will test this weekend.
            app = Application().start(r"C:\Program Files\PracticeApp\app.exe")
            time.sleep(3) # Wait slightly for older apps to load

            self.update_status("Connecting to Main Window...", color="yellow")
            window = app.main_window()
            
            self.update_status("Clicking 'File -> New Patient' (simulated)...", color="yellow")
            # In PyWinAuto, menu paths are separated by ->
            window.menu_select("File -> New Patient")
            
            self.update_status(f"Typing Data: {fname} {lname}...", color="yellow")
            
            # Type into specific legacy fields Suman identified
            window['First Name'].type_keys(fname)
            window['Last Name'].type_keys(lname)
            
            self.update_status("Clicking 'Save'...", color="yellow")
            window['Save'].click()

            self.update_status("Native UI Automation Complete! ✅", color="limegreen")
        except Exception as e:
            self.update_status(f"PyWinAuto Error: {e}", color="red")
        finally:
            self.enable_buttons()

    # ------------------ PyAutoGUI Logic (Screen Coordinate Fallback) ------------------
    def start_pyautogui_thread(self):
        self.disable_buttons()
        threading.Thread(target=self.run_pyautogui, daemon=True).start()

    def run_pyautogui(self):
        try:
            self.update_status("Launching App via Start Menu...", color="yellow")
            if platform.system() == "Windows":
                 pyautogui.hotkey('win', 'r')
                 time.sleep(1)
                 # pyautogui.typewrite(r'C:\Program Files\PracticeApp\app.exe')
                 pyautogui.typewrite('notepad.exe')
                 pyautogui.press('enter')
            elif platform.system() == "Darwin":
                 pyautogui.hotkey('command', 'space')
                 time.sleep(1)
                 pyautogui.typewrite('TextEdit')
                 pyautogui.press('enter')
                 
            self.update_status("Waiting 3s for app to load...", color="yellow")
            time.sleep(3) 

            self.update_status("Clicking arbitrary coordinates (x=450, y=300)...", color="yellow")
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
