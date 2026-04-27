@echo off
title Build Accessible Ebook Library Manager

echo This will build Accessible Ebook Library Manager.exe on this computer.
echo.
python --version >nul 2>&1
if errorlevel 1 (
  echo Python was not found.
  echo Install Python from https://www.python.org/downloads/windows/
  echo During install, check the box that says Add Python to PATH.
  pause
  exit /b 1
)

echo Installing PyInstaller if needed...
python -m pip install --upgrade pyinstaller
if errorlevel 1 (
  echo Could not install PyInstaller.
  pause
  exit /b 1
)

echo Building the EXE. This may take a few minutes.
python -m PyInstaller --onefile --windowed --name "Accessible Ebook Library Manager" library_manager.py
if errorlevel 1 (
  echo Build failed.
  pause
  exit /b 1
)

echo.
echo Done. Your EXE is in the dist folder:
echo dist\Accessible Ebook Library Manager.exe
echo.
pause
