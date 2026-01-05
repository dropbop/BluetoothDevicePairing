@echo off
REM Connect Bluetooth audio device
REM Edit "YOURDEVICE" below to your device name (partial match OK)

powershell -ExecutionPolicy Bypass -File "%~dp0bt-headphones.ps1" connect "YOURDEVICE"

timeout /t 3 >nul
