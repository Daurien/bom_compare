import sys
from cx_Freeze import setup, Executable

# Define the base
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="BOM_compare",
    version="0.1",
    description="compare two BOMs and returns the differences",
    executables=[Executable("interface.py", base=base)]
)
