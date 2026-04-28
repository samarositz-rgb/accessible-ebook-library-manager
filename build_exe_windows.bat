@echo off
title Build Accessible Ebook Library Manager

echo This will build Accessible Ebook Library Manager.exe on this computer.
echo.
set "APP_EXE=dist\Accessible Ebook Library Manager.exe"

echo Closing any running copy of Accessible Ebook Library Manager...
taskkill /IM "Accessible Ebook Library Manager.exe" /F >nul 2>&1
timeout /t 2 /nobreak >nul

if exist "%APP_EXE%" (
  del "%APP_EXE%" >nul 2>&1
  if exist "%APP_EXE%" (
    echo.
    echo The old EXE is still locked and cannot be replaced.
    echo Close Accessible Ebook Library Manager if it is running.
    echo If the dist folder is open in File Explorer, close that window too, then run this build again.
    echo.
    pause
    exit /b 1
  )
)

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
