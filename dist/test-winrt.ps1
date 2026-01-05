#Requires -Version 5.1
<#
.SYNOPSIS
    Test if WinRT types are accessible in PowerShell 5.1 without Add-Type
#>

Write-Host ""
Write-Host "=== WinRT Access Test ===" -ForegroundColor Cyan
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
Write-Host ""

$allPassed = $true

# Test 1: Load Windows.Foundation
Write-Host "[Test 1] Loading Windows.Foundation..." -ForegroundColor Yellow
try {
    [void][Windows.Foundation.IAsyncAction, Windows.Foundation, ContentType=WindowsRuntime]
    Write-Host "  PASSED - Windows.Foundation loaded" -ForegroundColor Green
} catch {
    Write-Host "  FAILED - $_" -ForegroundColor Red
    $allPassed = $false
}

# Test 2: Load Windows.Devices.Enumeration
Write-Host "[Test 2] Loading Windows.Devices.Enumeration..." -ForegroundColor Yellow
try {
    [void][Windows.Devices.Enumeration.DeviceInformation, Windows.Devices.Enumeration, ContentType=WindowsRuntime]
    Write-Host "  PASSED - DeviceInformation type loaded" -ForegroundColor Green
} catch {
    Write-Host "  FAILED - $_" -ForegroundColor Red
    $allPassed = $false
}

# Test 3: Load Windows.Devices.Bluetooth
Write-Host "[Test 3] Loading Windows.Devices.Bluetooth..." -ForegroundColor Yellow
try {
    [void][Windows.Devices.Bluetooth.BluetoothDevice, Windows.Devices.Bluetooth, ContentType=WindowsRuntime]
    Write-Host "  PASSED - BluetoothDevice type loaded" -ForegroundColor Green
} catch {
    Write-Host "  FAILED - $_" -ForegroundColor Red
    $allPassed = $false
}

# Test 4: Try to enumerate Bluetooth devices
Write-Host "[Test 4] Enumerating Bluetooth devices..." -ForegroundColor Yellow
try {
    # Get the device selector for paired Bluetooth devices
    $selector = [Windows.Devices.Bluetooth.BluetoothDevice]::GetDeviceSelectorFromPairingState($true)
    Write-Host "  Selector: $selector" -ForegroundColor Gray

    # Find devices (async)
    $asyncOp = [Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($selector)

    # Wait for completion (with timeout)
    $timeout = 10
    $elapsed = 0
    while ($asyncOp.Status -eq 'Started' -and $elapsed -lt $timeout) {
        Start-Sleep -Milliseconds 500
        $elapsed += 0.5
    }

    if ($asyncOp.Status -eq 'Completed') {
        $devices = $asyncOp.GetResults()
        Write-Host "  Found $($devices.Count) paired Bluetooth device(s):" -ForegroundColor Green

        foreach ($device in $devices) {
            Write-Host "    - $($device.Name)" -ForegroundColor White
        }
    } else {
        Write-Host "  Status: $($asyncOp.Status)" -ForegroundColor Yellow
        $allPassed = $false
    }
} catch {
    Write-Host "  FAILED - $_" -ForegroundColor Red
    $allPassed = $false
}

# Summary
Write-Host ""
Write-Host "=== Summary ===" -ForegroundColor Cyan
if ($allPassed) {
    Write-Host "SUCCESS: WinRT types are accessible!" -ForegroundColor Green
    Write-Host "We can use WinRT APIs for Bluetooth control." -ForegroundColor Green
} else {
    Write-Host "PARTIAL/FAILED: Some WinRT features blocked." -ForegroundColor Yellow
    Write-Host "Will need to use Settings UI automation instead." -ForegroundColor Yellow
}
Write-Host ""
