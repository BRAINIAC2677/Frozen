from cx_Freeze import setup, Executable
import sys

base = None

base = None
if (sys.platform == "win32"):
    base = "Win32GUI"

executables = [Executable("Compassion.py", base=base)]

packages = ["idna", "win10toast", 'threading', 'playsound', 'datetime', 'openpyxl', 'tkinter', 'sys']
options = {
    'build_exe': {
        'packages':packages, 'include_files': [r'logo.ico', r'VoiceMsg.wav']
    },
}

setup(
    name = "Frozen",
    options = options,
    version = "1.0",
    description = 'Frozen',
    executables = executables
)