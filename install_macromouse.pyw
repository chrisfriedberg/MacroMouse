import subprocess
import sys
import os
from pathlib import Path

BASE_DIR = Path(__file__).parent
os.chdir(BASE_DIR)

def install_dependencies():
    # installs runtime deps plus PyInstaller
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "pyperclip", "pyinstaller"])

def build_exe():
    # build MacroMouse.exe without console, include macros and log files
    subprocess.check_call([
        sys.executable, "-m", "PyInstaller",
        "--noconsole",
        "--name", "MacroMouse",
        "--add-data", f"{BASE_DIR / 'macros.txt'};.",
        "--add-data", f"{BASE_DIR / 'MacroMouse.log'};.",
        str(BASE_DIR / "MacroMouse.pyw")
    ])

if __name__ == "__main__":
    install_dependencies()
    build_exe()
    print("✅ Build complete. Look in dist\MacroMouse for MacroMouse.exe")
