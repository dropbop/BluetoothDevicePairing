# Bluetooth Audio Connect/Disconnect

Minimal scripts to connect and disconnect Bluetooth audio devices on locked-down Windows systems.

## Quick Start

### 1. Find your device name

```cmd
powershell -ExecutionPolicy Bypass -File bt-headphones.ps1 list
```

Example output:
```
Bluetooth Audio Devices:
--------------------------------------------------
Sony WH-1000XM5 [Connected]
AirPods Pro [Disconnected]
```

### 2. Edit the batch files

Open `bt-on.bat` and `bt-off.bat` in Notepad.

Replace `YOURDEVICE` with your device name (or part of it):

```batch
powershell -ExecutionPolicy Bypass -File "%~dp0bt-headphones.ps1" connect "WH-1000XM5"
```

Partial names work - `"XM5"` or `"Sony"` would also match `"Sony WH-1000XM5"`.

### 3. Test

Double-click `bt-on.bat` to connect, `bt-off.bat` to disconnect.

### 4. Pin to Taskbar (Optional)

1. Right-click the `.bat` file
2. Select "Create shortcut"
3. Right-click the shortcut → "Pin to taskbar"

## Usage

### List devices
```
.\bt-headphones.ps1 list
```

### Connect
```
.\bt-headphones.ps1 connect "DeviceName"
```

### Disconnect
```
.\bt-headphones.ps1 disconnect "DeviceName"
```

## Requirements

- Windows 10 1809+ or Windows 11
- Device must be **paired first** via Windows Settings > Bluetooth
- PowerShell (built-in)

## Troubleshooting

### "No Bluetooth audio devices found"
- Make sure the device is paired (via Windows Settings > Bluetooth)
- Turn on the device and make sure it's in range
- Try running `list` again after a few seconds

### "Multiple devices match"
- Use a more specific name to avoid ambiguity
- Example: Use `"WH-1000XM5"` instead of just `"Sony"`

### Connection takes time
- The script sends a connect/disconnect request to Windows
- The actual connection may take 2-5 seconds
- Some devices take longer on first connection

### Script fails silently
- Make sure you're running from the correct directory
- Try running the PowerShell script directly to see errors:
  ```
  powershell -ExecutionPolicy Bypass -File bt-headphones.ps1 list
  ```

## How It Works

The script uses Windows Core Audio APIs to:
1. Enumerate all audio endpoints
2. Filter for Bluetooth devices (device path starts with `{2}.\\?\bth`)
3. Send Kernel Streaming property requests to connect/disconnect

No external dependencies - everything is inline C# compiled at runtime.
