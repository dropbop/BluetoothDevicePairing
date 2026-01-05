<#
.SYNOPSIS
    Connect/disconnect Bluetooth audio devices using PnP cmdlets (no Add-Type required)

.DESCRIPTION
    Uses built-in Windows PnP cmdlets to enable/disable Bluetooth audio devices.
    Disabling = disconnect, Enabling = reconnect.

.EXAMPLE
    .\bt-pnp.ps1 list
    .\bt-pnp.ps1 connect "WH-1000XM5"
    .\bt-pnp.ps1 disconnect "WH-1000XM5"
#>

param(
    [Parameter(Position=0)]
    [string]$Action = "help",

    [Parameter(Position=1)]
    [string]$DeviceName = ""
)

$ErrorActionPreference = "Stop"

function Get-BluetoothAudioDevices {
    # Get Bluetooth audio devices - they show up as "Bluetooth" class or have BTH in hardware IDs
    # Also check for audio-related Bluetooth devices

    $devices = @()

    # Method 1: Look for Bluetooth devices
    try {
        $btDevices = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue |
            Where-Object { $_.FriendlyName -and $_.FriendlyName -notmatch "Radio|Adapter|Enumerator" }

        foreach ($dev in $btDevices) {
            $devices += [PSCustomObject]@{
                Name = $dev.FriendlyName
                InstanceId = $dev.InstanceId
                Status = $dev.Status
                Class = $dev.Class
            }
        }
    } catch {}

    # Method 2: Look for audio devices with Bluetooth in the name or hardware path
    try {
        $audioDevices = Get-PnpDevice -Class AudioEndpoint -ErrorAction SilentlyContinue |
            Where-Object { $_.FriendlyName }

        foreach ($dev in $audioDevices) {
            # Check if it's a Bluetooth device by looking at the instance ID
            if ($dev.InstanceId -match "BTH|BTHENUM") {
                $devices += [PSCustomObject]@{
                    Name = $dev.FriendlyName
                    InstanceId = $dev.InstanceId
                    Status = $dev.Status
                    Class = $dev.Class
                }
            }
        }
    } catch {}

    # Method 3: Look in Media class
    try {
        $mediaDevices = Get-PnpDevice -Class Media -ErrorAction SilentlyContinue |
            Where-Object { $_.InstanceId -match "BTH|BTHENUM" -and $_.FriendlyName }

        foreach ($dev in $mediaDevices) {
            $devices += [PSCustomObject]@{
                Name = $dev.FriendlyName
                InstanceId = $dev.InstanceId
                Status = $dev.Status
                Class = $dev.Class
            }
        }
    } catch {}

    # Method 4: Sound class
    try {
        $soundDevices = Get-PnpDevice -Class Sound -ErrorAction SilentlyContinue |
            Where-Object { $_.InstanceId -match "BTH|BTHENUM" -and $_.FriendlyName }

        foreach ($dev in $soundDevices) {
            $devices += [PSCustomObject]@{
                Name = $dev.FriendlyName
                InstanceId = $dev.InstanceId
                Status = $dev.Status
                Class = $dev.Class
            }
        }
    } catch {}

    # Remove duplicates by InstanceId
    $devices = $devices | Sort-Object InstanceId -Unique

    return $devices
}

function Show-Help {
    Write-Host ""
    Write-Host "Bluetooth Audio Device Manager (PnP version)" -ForegroundColor Cyan
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  .\bt-pnp.ps1 list                    List Bluetooth devices"
    Write-Host "  .\bt-pnp.ps1 connect <name>          Enable/reconnect device"
    Write-Host "  .\bt-pnp.ps1 disconnect <name>       Disable/disconnect device"
    Write-Host "  .\bt-pnp.ps1 help                    Show this help"
    Write-Host ""
    Write-Host "Note: connect/disconnect may require admin rights." -ForegroundColor Gray
    Write-Host ""
}

function Show-List {
    Write-Host ""
    Write-Host "Scanning for Bluetooth devices..." -ForegroundColor Cyan
    Write-Host ""

    $devices = Get-BluetoothAudioDevices

    if ($devices.Count -eq 0) {
        Write-Host "No Bluetooth audio devices found." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Make sure your device is:" -ForegroundColor Gray
        Write-Host "  - Paired with this computer" -ForegroundColor Gray
        Write-Host "  - Turned on" -ForegroundColor Gray
        Write-Host ""

        # Show all Bluetooth devices for debugging
        Write-Host "All Bluetooth class devices:" -ForegroundColor Gray
        Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue |
            Format-Table FriendlyName, Status, InstanceId -AutoSize
        return
    }

    Write-Host "Bluetooth Devices:" -ForegroundColor Green
    Write-Host ("-" * 60)

    $index = 1
    foreach ($dev in $devices) {
        $statusColor = if ($dev.Status -eq "OK") { "Green" } else { "Yellow" }
        Write-Host ("  [{0}] {1}" -f $index, $dev.Name) -ForegroundColor White -NoNewline
        Write-Host (" [{0}]" -f $dev.Status) -ForegroundColor $statusColor
        Write-Host ("      Class: {0}" -f $dev.Class) -ForegroundColor Gray
        $index++
    }

    Write-Host ""
    Write-Host "Tip: Use 'disconnect' or 'connect' with '--all' to target all" -ForegroundColor Gray
    Write-Host "     devices with the same name at once." -ForegroundColor Gray
    Write-Host ""
}

function Set-DeviceState {
    param(
        [string]$Name,
        [bool]$Enable
    )

    $action = if ($Enable) { "connect" } else { "disconnect" }
    $actionVerb = if ($Enable) { "Enabling" } else { "Disabling" }
    $actionPast = if ($Enable) { "enabled" } else { "disabled" }

    if ([string]::IsNullOrWhiteSpace($Name)) {
        Write-Host "Error: Device name required." -ForegroundColor Red
        Write-Host "Usage: .\bt-pnp.ps1 $action ""DeviceName""" -ForegroundColor Yellow
        return
    }

    $devices = Get-BluetoothAudioDevices

    # Find matching devices (partial, case-insensitive)
    $matches = @($devices | Where-Object { $_.Name -like "*$Name*" })

    if ($matches.Count -eq 0) {
        Write-Host "No device found matching: $Name" -ForegroundColor Red
        Write-Host ""
        Write-Host "Available devices:" -ForegroundColor Yellow
        $uniqueNames = $devices | Select-Object -ExpandProperty Name -Unique
        foreach ($n in $uniqueNames) {
            Write-Host "  - $n"
        }
        return
    }

    # Multiple matches is OK - Bluetooth devices have multiple entries (audio, AVRCP, etc.)
    # We'll target all of them since they're part of the same physical device
    Write-Host ""
    Write-Host "Found $($matches.Count) device entries matching '$Name':" -ForegroundColor Cyan
    foreach ($dev in $matches) {
        Write-Host "  - $($dev.Name) [$($dev.Class)] - $($dev.Status)" -ForegroundColor Gray
    }
    Write-Host ""

    $successCount = 0
    $failCount = 0
    $skipCount = 0

    foreach ($device in $matches) {
        try {
            Write-Host "$actionVerb $($device.Name) [$($dev.Class)]..." -ForegroundColor Yellow -NoNewline

            if ($Enable) {
                if ($device.Status -eq "OK") {
                    Write-Host " already enabled, skipping" -ForegroundColor Gray
                    $skipCount++
                    continue
                }
                Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction Stop
            } else {
                if ($device.Status -ne "OK") {
                    Write-Host " already disabled, skipping" -ForegroundColor Gray
                    $skipCount++
                    continue
                }
                Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction Stop
            }

            Write-Host " OK" -ForegroundColor Green
            $successCount++

        } catch {
            Write-Host " FAILED" -ForegroundColor Red
            $failCount++

            if ($_.Exception.Message -match "Access|denied|admin|privilege") {
                Write-Host "  Requires administrator privileges" -ForegroundColor Yellow
            } else {
                Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    Write-Host ""
    if ($successCount -gt 0) {
        Write-Host "$successCount device(s) $actionPast successfully." -ForegroundColor Green
    }
    if ($skipCount -gt 0) {
        Write-Host "$skipCount device(s) skipped (already in desired state)." -ForegroundColor Gray
    }
    if ($failCount -gt 0) {
        Write-Host "$failCount device(s) failed." -ForegroundColor Red
        Write-Host ""
        Write-Host "Try running as Administrator:" -ForegroundColor Yellow
        Write-Host "  Right-click the .bat file -> Run as administrator" -ForegroundColor White
    }

    if ($successCount -gt 0) {
        Write-Host ""
        Write-Host "Note: Audio switching may take a few seconds." -ForegroundColor Gray
    }
}

# Main
switch ($Action.ToLower()) {
    "list" { Show-List }
    "connect" { Set-DeviceState -Name $DeviceName -Enable $true }
    "disconnect" { Set-DeviceState -Name $DeviceName -Enable $false }
    "help" { Show-Help }
    default { Show-Help }
}

Write-Host ""
