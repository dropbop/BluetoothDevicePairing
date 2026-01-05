# Add-Type Compatibility Test

This test verifies that PowerShell's `Add-Type` with inline C# works on your locked-down system.

## How to Run

### Option 1: Double-click (Recommended)
1. Double-click `test-addtype.bat`
2. A command window will open and run the test
3. Check the output for SUCCESS or FAILED

### Option 2: Command Line
```cmd
cd test
test-addtype.bat
```

### Option 3: PowerShell directly
```powershell
powershell -ExecutionPolicy Bypass -File test-addtype.ps1
```

## What It Tests

1. **Basic Add-Type** - Can PowerShell compile and run inline C# code?
2. **InteropServices** - Can we use COM attributes and struct layouts?
3. **COM Objects** - Can we create COM objects?

## Expected Output (Success)

```
Testing PowerShell Add-Type with inline C#...

[Test 1] Basic inline C# class...
  Result: Hello from inline C#!
  System: CLR: 4.0.30319.42000, OS: Win32NT, 64-bit: True
  PASSED

[Test 2] InteropServices and COM attributes...
  Result: Struct created - GUID: ..., Value: 42
  PASSED

[Test 3] COM object creation...
  Shell.Application COM object created successfully
  PASSED

========================================
SUCCESS: All tests passed!
Add-Type with inline C# works on this system.
The Bluetooth audio script should work.
========================================
```

## If It Fails

If you see a FAILED message, the Bluetooth script won't work on your system.
Common reasons for failure:
- **Constrained Language Mode** - PowerShell is restricted to safe cmdlets only
- **AppLocker/WDAC policies** - Code compilation is blocked
- **Antivirus blocking** - Security software may block Add-Type

Please share the error message so we can explore alternatives.
