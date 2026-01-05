#Requires -Version 5.1
<#
.SYNOPSIS
    Test WinRT async with different patterns
#>

Write-Host ""
Write-Host "=== WinRT Async Test v2 ===" -ForegroundColor Cyan
Write-Host ""

# Helper to await async operations
function Await-Operation {
    param($AsyncOp, $TimeoutSeconds = 15)

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if ($AsyncOp.Status -eq [Windows.Foundation.AsyncStatus]::Completed) {
            return $AsyncOp.GetResults()
        }
        if ($AsyncOp.Status -eq [Windows.Foundation.AsyncStatus]::Error) {
            throw "Async operation failed: $($AsyncOp.ErrorCode)"
        }
        if ($AsyncOp.Status -eq [Windows.Foundation.AsyncStatus]::Canceled) {
            throw "Async operation canceled"
        }
        Start-Sleep -Milliseconds 100
    }

    throw "Timeout after $TimeoutSeconds seconds (Status: $($AsyncOp.Status))"
}

# Load required types
Write-Host "Loading WinRT types..." -ForegroundColor Yellow
try {
    # Load in specific order
    $null = [Windows.Foundation.AsyncStatus, Windows.Foundation, ContentType=WindowsRuntime]
    $null = [Windows.Devices.Enumeration.DeviceInformation, Windows.Devices.Enumeration, ContentType=WindowsRuntime]
    $null = [Windows.Devices.Bluetooth.BluetoothDevice, Windows.Devices.Bluetooth, ContentType=WindowsRuntime]
    Write-Host "  Types loaded OK" -ForegroundColor Green
} catch {
    Write-Host "  FAILED: $_" -ForegroundColor Red
    exit 1
}

# Method 1: Use BluetoothDevice selector
Write-Host ""
Write-Host "[Method 1] BluetoothDevice.GetDeviceSelectorFromPairingState..." -ForegroundColor Yellow
try {
    $selector = [Windows.Devices.Bluetooth.BluetoothDevice]::GetDeviceSelectorFromPairingState($true)
    Write-Host "  Selector obtained" -ForegroundColor Gray

    $asyncOp = [Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($selector)
    Write-Host "  Async operation started, waiting..." -ForegroundColor Gray

    $devices = Await-Operation $asyncOp -TimeoutSeconds 15
    Write-Host "  SUCCESS: Found $($devices.Count) device(s)" -ForegroundColor Green

    foreach ($d in $devices) {
        Write-Host "    - $($d.Name) [$($d.Id)]" -ForegroundColor White
    }
} catch {
    Write-Host "  FAILED: $_" -ForegroundColor Red
}

# Method 2: Simpler AQS query
Write-Host ""
Write-Host "[Method 2] Simple Bluetooth AQS query..." -ForegroundColor Yellow
try {
    # Simpler query for Bluetooth devices
    $aqsFilter = "System.Devices.Aep.ProtocolId:=`"{e0cbf06c-cd8b-4647-bb8a-263b43f0f974}`""

    $asyncOp = [Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($aqsFilter, $null, [Windows.Devices.Enumeration.DeviceInformationKind]::AssociationEndpoint)
    Write-Host "  Async operation started, waiting..." -ForegroundColor Gray

    $devices = Await-Operation $asyncOp -TimeoutSeconds 15
    Write-Host "  SUCCESS: Found $($devices.Count) device(s)" -ForegroundColor Green

    foreach ($d in $devices) {
        Write-Host "    - $($d.Name)" -ForegroundColor White
    }
} catch {
    Write-Host "  FAILED: $_" -ForegroundColor Red
}

# Method 3: Try to get a specific device by creating BluetoothDevice directly
Write-Host ""
Write-Host "[Method 3] Check if BluetoothDevice.FromIdAsync works..." -ForegroundColor Yellow
Write-Host "  (This would be used after finding device ID)" -ForegroundColor Gray
Write-Host "  Skipping - need device ID first" -ForegroundColor Gray

Write-Host ""
Write-Host "=== Test Complete ===" -ForegroundColor Cyan
Write-Host ""
