from cx_Freeze import setup, Executable
import os

base = None

executables = [Executable("Monitor.py", base=base)]

packages = ["idna", "os", "time", "datetime", "watchdog", "ctypes", "collections", "threading",
            "colorama", "argh", "pathtools", "sys", "pandas", "numpy", "openpyxl", "PIL"]
options = {
    'build_exe': {
        'packages': packages,
        "include_msvcr": True,
        "include_files": ['config.ini', 'assays.txt', 'Formulas.xlsx'],
        'excludes': ['tkinter']
                  },
          }

setup(
    name="LabHelper",
    options=options,
    version="1.0",
    description="Tool for T121 Genotypers at Sanger",
    executables=executables, requires=['colorama', 'watchdog', 'pandas', 'numpy', 'pillow']
)
