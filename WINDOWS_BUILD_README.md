# Document Generator Pro - Build Instructions

To convert this Python application into a standalone Windows `.exe` file that users can run without installing Python, follow these instructions. 

This MUST be executed on a Windows operating system.

## Step 1: Install Build Requirements
Open your Command Prompt (cmd.exe) or PowerShell and run:
```cmd
pip install customtkinter pyinstaller python-docx packaging
```

## Step 2: Build the Executable
In the same folder as `app.py`, run this PyInstaller command:

```cmd
pyinstaller --noconsole --onefile --windowed --name "DocumentGeneratorPro" app.py
```

### What these flags mean:
* `--noconsole` / `--windowed`: Hides the black CMD window that normally opens when you launch a Python script. Essential for a GUI app.
* `--onefile`: Bundles everything (your code, python, docx library, UI libraries) into a single, clean `.exe` file instead of a messy folder.
* `--name`: What the final `.exe` file will be named.

## Step 3: Distribute!
1. PyInstaller will create a new folder called `dist`.
2. Inside `dist`, you will find your shiny new `DocumentGeneratorPro.exe` file.
3. You can copy this `.exe` file, zip it up, and sell/distribute it! Anyone with Windows 10/11 can double click it and it will just work!
