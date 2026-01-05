@echo off
REM Attempt to disconnect Bluetooth device via Settings UI automation
REM Edit "OpenRun" to match your device name

echo.
echo Attempting to disconnect Bluetooth device...
echo DO NOT touch keyboard/mouse until complete.
echo.

powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0bt-settings.ps1" disconnect "OpenRun"

pause
