@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build Application.ps1"
if errorlevel 1 (
  echo Build failed. The previous application has been kept.
  pause
  exit /b 1
)
pause
