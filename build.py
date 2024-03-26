import cx_Freeze
import sys
import openpyxl
import PIL
import gtts
import tkinter
import time
import serial

base = None

if sys.platform == 'win32':
    base = "Win32GUI"

executables = [cx_Freeze.Executable("main.py", base=base, icon="Logo.png")]

cx_Freeze.setup(
    name = "ShopSense-Connect",
    options = {"build_exe": {"packages":["tkinter", "serial" , "openpyxl" , "time", "PIL", "gtts"], "include_files":["Logo.png","RIFD.png"]}},
    version = "2024.1",
    description = "Manage ShopSense Data",
    executables = executables
    )