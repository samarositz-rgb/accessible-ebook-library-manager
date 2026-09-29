import sys
import ctypes
from ctypes import wintypes
from pathlib import Path
import os

ctypes.windll.kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
ctypes.windll.kernel32.GetModuleHandleW.restype = wintypes.HMODULE
ctypes.windll.kernel32.GetModuleFileNameW.argtypes = [wintypes.HMODULE, wintypes.LPWSTR, wintypes.DWORD]
for name in ("VCRUNTIME140.dll", "MSVCP140.dll", "Qt6Core.dll"):
    handle = ctypes.windll.kernel32.GetModuleHandleW(name)
    if handle:
        path = ctypes.create_unicode_buffer(32768)
        ctypes.windll.kernel32.GetModuleFileNameW(handle, path, len(path))
        print(f"Loaded {name}: {path.value}", flush=True)

root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
for directory in (root / "PySide6", root / "shiboken6"):
    if directory.is_dir():
        os.add_dll_directory(str(directory))
for relative in ("PySide6/Qt6Core.dll", "shiboken6/shiboken6.abi3.dll", "PySide6/pyside6.abi3.dll"):
    path = root / relative
    try:
        ctypes.WinDLL(str(path))
        print(f"Loaded {relative}", flush=True)
    except OSError as exc:
        print(f"Failed {relative}: {exc}", flush=True)

from PySide6.QtCore import qVersion

print(f"Python {sys.version.split()[0]}; Qt {qVersion()}", flush=True)
