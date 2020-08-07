#from distutils.core import setup
#import py2exe
#
#setup(windows=['liveppt.py'])


import sys
from cx_Freeze import setup, Executable


# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"], 
					 #"includes": ["newhymal.hdb", "outline.bas"],
                     "excludes": ["numpy", "matplotlib", "tkinter"],
					 'build_exe': 'LPPT'
					 }

exe = Executable("liveppt.py", base = base, icon='lppt.ico')

setup(  name = "LibPPT",
        version = "0.1",
        description = "Convert a hymal ppt to a subtitle ppt",
        options = {"build_exe": build_exe_options},
        executables = [exe])