import customtkinter as ctk
import pyautogui
import time
import threading
import os
import platform
import subprocess

# Set the appearance and color theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class AutomationAgent(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title("Visible Word Automation Bot")
        self.geometry(f"{600}x{400}")
        self.resizable(False, False)

        # ------------------ UI Layout ------------------
        # Header Label
        self.header_label = ctk.CTkLabel(self, text="🤖 Live Bot Automation", font=ctk.CTkFont(size=24, weight="bold"))
        self.header_label.pack(pady=(20, 10))

        # Instructions
        self.instructions = ctk.CTkLabel(self, text="This agent will take control of your mouse and keyboard\nto visually open Word and type the text below.", text_color="gray")
        self.instructions.pack(pady=(0, 10))

        # Text Input
        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.pack(pady=10, padx=20, fill="x")
        
        self.doc_text_label = ctk.CTkLabel(self.input_frame, text="Text to Type:")
        self.doc_text_label.pack(side="left", padx=10, pady=10)
        
        self.doc_text_entry = ctk.CTkEntry(self.input_frame, placeholder_text="Type something here...", width=350)
        self.doc_text_entry.pack(side="left", padx=10, pady=10, fill="x", expand=True)
        self.doc_text_entry.insert(0, "gey")

        # Status Display
        self.status_frame = ctk.CTkFrame(self, fg_color="black")
        self.status_frame.pack(pady=10, padx=20, fill="both", expand=True)

        self.status_label = ctk.CTkLabel(self.status_frame, text="Status: Idle (Waiting for command)", font=ctk.CTkFont(size=14), text_color="limegreen")
        self.status_label.pack(pady=30, padx=10)

        # Action Buttons
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(pady=(10, 20), padx=20, fill="x")

        self.start_btn = ctk.CTkButton(
            self.button_frame, 
            text="🚀 Start Live Automation", 
            font=ctk.CTkFont(size=16, weight="bold"),
            height=45,
            command=self.start_automation_safely
        )
        self.start_btn.pack(side="bottom", expand=True, fill="x", padx=10)

    # ------------------ Action Logic ------------------
    def update_status(self, text, color="limegreen"):
        # We use 'after' to safely update Tkinter UI from a background thread
        self.after(0, lambda: self.status_label.configure(text=f"Status: {text}", text_color=color))

    def enable_button(self, text="🚀 Start Live Automation"):
        self.after(0, lambda: self.start_btn.configure(state="normal", text=text))

    def start_automation_safely(self):
        text_to_type = self.doc_text_entry.get()
        if not text_to_type.strip():
             self.update_status("Error: Please provide text to type.", color="red")
             return

        # Disable button to prevent multiple clicks
        self.start_btn.configure(state="disabled", text="Running...")
        
        # Run automation in a separate thread so UI doesn't freeze!
        threading.Thread(target=self.run_bot_logic, args=(text_to_type,), daemon=True).start()

    def run_bot_logic(self, text_to_type):
        try:
            self.update_status("Step 1: Opening Microsoft Word...", color="yellow")
            
            # Cross-platform Word Launching
            if platform.system() == "Windows":
                 # Open Word via Windows Terminal
                 os.system("start winword")
                 
                 self.update_status("Step 2: Waiting for Word to load...", color="yellow")
                 time.sleep(6) # Wait for splash screen

                 self.update_status("Step 3: Creating Blank Document...", color="yellow")
                 pyautogui.press('enter') # Bypasses 'Home' screen to make a blank doc
                 time.sleep(2)
                 
            elif platform.system() == "Darwin": # Mac
                 # Open Word via Terminal
                 subprocess.Popen(["open", "-a", "Microsoft Word"])
                 
                 self.update_status("Step 2: Waiting for Word to load...", color="yellow")
                 time.sleep(4)
                 
                 self.update_status("Step 3: Creating Blank Document...", color="yellow")
                 pyautogui.hotkey('command', 'n') # Opens new doc from gallery
                 time.sleep(2)
                 
            else:
                 self.update_status("OS Not Supported. Try Windows or Mac.", color="red")
                 self.enable_button()
                 return
                 
            self.update_status("Step 4: Executing visual typing...", color="cyan")
            
            # The magic: physically typing it keystroke by keystroke
            # `interval` determines the speed of the typing (e.g. 0.1 secs between strokes)
            pyautogui.write(text_to_type, interval=0.15)

            # Re-enable button
            self.update_status("Finished: Automation Sequence Complete! ✨", color="limegreen")
            self.enable_button()

        except Exception as e:
            self.update_status(f"Error during sequence:\n{e}", color="red")
            self.enable_button()

if __name__ == "__main__":
    # Safety feature for PyAutoGUI (Move mouse rapidly to any corner of screen to Force Abort)
    pyautogui.FAILSAFE = True
    app = AutomationAgent()
    app.mainloop()
