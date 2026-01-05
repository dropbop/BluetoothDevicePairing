#Requires -Version 5.1
<#
.SYNOPSIS
    Connect/disconnect Bluetooth devices via Settings UI automation

.DESCRIPTION
    Opens Windows 11 Bluetooth settings and uses keyboard navigation
    to connect or disconnect a device. Works without admin rights.

.PARAMETER Action
    connect, disconnect, or open (just opens settings)

.PARAMETER DeviceName
    Name of the Bluetooth device (used for visual confirmation)

.EXAMPLE
    .\bt-settings.ps1 open
    .\bt-settings.ps1 connect "OpenRun"
    .\bt-settings.ps1 disconnect "OpenRun"
#>

param(
    [Parameter(Position=0)]
    [ValidateSet("connect", "disconnect", "open", "help")]
    [string]$Action = "help",

    [Parameter(Position=1)]
    [string]$DeviceName = "OpenRun"
)

$ErrorActionPreference = "Stop"

function Show-Help {
    Write-Host ""
    Write-Host "Bluetooth Settings Automation" -ForegroundColor Cyan
    Write-Host "=============================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  .\bt-settings.ps1 open                    Open Bluetooth settings"
    Write-Host "  .\bt-settings.ps1 connect [DeviceName]    Connect device"
    Write-Host "  .\bt-settings.ps1 disconnect [DeviceName] Disconnect device"
    Write-Host ""
    Write-Host "How it works:" -ForegroundColor Yellow
    Write-Host "  1. Opens Windows Settings > Bluetooth"
    Write-Host "  2. Uses keyboard navigation to find your device"
    Write-Host "  3. Triggers Connect or Disconnect"
    Write-Host ""
    Write-Host "Note:" -ForegroundColor Yellow
    Write-Host "  - Don't touch keyboard/mouse while script runs"
    Write-Host "  - Device must be visible in Bluetooth settings"
    Write-Host "  - May need adjustment for your specific setup"
    Write-Host ""
}

function Open-BluetoothSettings {
    Write-Host "Opening Bluetooth settings..." -ForegroundColor Cyan

    # Close any existing Settings window first
    Get-Process "SystemSettings" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500

    # Open Bluetooth & devices settings
    Start-Process "ms-settings:bluetooth"
    Start-Sleep -Seconds 2

    # Ensure Settings window is focused
    $shell = New-Object -ComObject WScript.Shell
    $focused = $shell.AppActivate("Settings")

    if (-not $focused) {
        Start-Sleep -Seconds 1
        $focused = $shell.AppActivate("Settings")
    }

    return $shell
}

function Send-Keys {
    param($Shell, $Keys, $Delay = 300)
    $Shell.SendKeys($Keys)
    Start-Sleep -Milliseconds $Delay
}

function Connect-Device {
    param([string]$Name)

    Write-Host ""
    Write-Host "Attempting to CONNECT: $Name" -ForegroundColor Green
    Write-Host ""
    Write-Host "Please don't touch keyboard/mouse..." -ForegroundColor Yellow
    Write-Host ""

    $shell = Open-BluetoothSettings
    Start-Sleep -Seconds 1

    Write-Host "Navigating to device list..." -ForegroundColor Gray

    # Windows 11 Bluetooth settings layout:
    # - Tab navigates through elements
    # - Devices are in a list
    # - Enter on a device expands it
    # - Connect button appears when expanded

    # Strategy: Tab into the device list area, then search
    # This is approximate and may need adjustment

    # Tab to get into the main content area (skip sidebar)
    Send-Keys $shell "{TAB}" 200
    Send-Keys $shell "{TAB}" 200
    Send-Keys $shell "{TAB}" 200
    Send-Keys $shell "{TAB}" 200

    Write-Host "Looking for device in list..." -ForegroundColor Gray

    # Try to find the device by tabbing through the list
    # Each device is a clickable item; Enter expands it
    # We'll tab through and hope to land on the device

    # Tab through potential device entries (up to 10)
    for ($i = 0; $i -lt 10; $i++) {
        Send-Keys $shell "{TAB}" 150
    }

    # Now try pressing Enter to expand a device, then look for Connect
    Write-Host "Attempting to expand device and connect..." -ForegroundColor Gray
    Send-Keys $shell "{ENTER}" 500

    # After expanding, Tab to Connect button and press Enter
    Send-Keys $shell "{TAB}" 200
    Send-Keys $shell "{TAB}" 200
    Send-Keys $shell "{ENTER}" 300

    Write-Host ""
    Write-Host "Command sent. Check if device connected." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "If this didn't work:" -ForegroundColor Yellow
    Write-Host "  1. Run '.\bt-settings.ps1 open' to open settings"
    Write-Host "  2. Manually click on '$Name'"
    Write-Host "  3. Click 'Connect'"
    Write-Host ""

    # Close Settings after a delay
    Start-Sleep -Seconds 2
    Get-Process "SystemSettings" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Disconnect-Device {
    param([string]$Name)

    Write-Host ""
    Write-Host "Attempting to DISCONNECT: $Name" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please don't touch keyboard/mouse..." -ForegroundColor Yellow
    Write-Host ""

    $shell = Open-BluetoothSettings
    Start-Sleep -Seconds 1

    Write-Host "Navigating to device list..." -ForegroundColor Gray

    # Similar navigation as Connect
    Send-Keys $shell "{TAB}" 200
    Send-Keys $shell "{TAB}" 200
    Send-Keys $shell "{TAB}" 200
    Send-Keys $shell "{TAB}" 200

    Write-Host "Looking for device in list..." -ForegroundColor Gray

    for ($i = 0; $i -lt 10; $i++) {
        Send-Keys $shell "{TAB}" 150
    }

    Write-Host "Attempting to expand device and disconnect..." -ForegroundColor Gray
    Send-Keys $shell "{ENTER}" 500

    # Disconnect button is usually in similar position to Connect
    Send-Keys $shell "{TAB}" 200
    Send-Keys $shell "{TAB}" 200
    Send-Keys $shell "{ENTER}" 300

    Write-Host ""
    Write-Host "Command sent. Check if device disconnected." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "If this didn't work:" -ForegroundColor Yellow
    Write-Host "  1. Run '.\bt-settings.ps1 open' to open settings"
    Write-Host "  2. Manually click on '$Name'"
    Write-Host "  3. Click 'Disconnect'"
    Write-Host ""

    Start-Sleep -Seconds 2
    Get-Process "SystemSettings" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Open-Only {
    Write-Host ""
    Write-Host "Opening Bluetooth settings..." -ForegroundColor Cyan
    Write-Host "Manually connect/disconnect your device from the UI." -ForegroundColor Gray
    Write-Host ""

    Start-Process "ms-settings:bluetooth"
}

# Main
switch ($Action.ToLower()) {
    "connect"    { Connect-Device -Name $DeviceName }
    "disconnect" { Disconnect-Device -Name $DeviceName }
    "open"       { Open-Only }
    "help"       { Show-Help }
    default      { Show-Help }
}
