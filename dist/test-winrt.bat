@echo off
REM Test WinRT access in PowerShell 5.1
REM IMPORTANT: Must use Windows PowerShell (powershell.exe), not PowerShell 7 (pwsh.exe)

powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0test-winrt.ps1"

echo.
pause
