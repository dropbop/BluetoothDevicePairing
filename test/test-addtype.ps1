# Test script to verify Add-Type with inline C# works on your system
# This tests the same mechanism the Bluetooth script will use

$ErrorActionPreference = "Stop"

Write-Host "Testing PowerShell Add-Type with inline C#..." -ForegroundColor Cyan
Write-Host ""

try {
    # Test 1: Basic Add-Type with simple C# class
    Write-Host "[Test 1] Basic inline C# class..." -ForegroundColor Yellow

    $csharpCode = @"
using System;

public class TestClass
{
    public static string GetMessage()
    {
        return "Hello from inline C#!";
    }

    public static string GetSystemInfo()
    {
        return string.Format(
            "CLR: {0}, OS: {1}, 64-bit: {2}",
            Environment.Version,
            Environment.OSVersion.Platform,
            Environment.Is64BitProcess
        );
    }
}
"@

    Add-Type -TypeDefinition $csharpCode -Language CSharp

    $message = [TestClass]::GetMessage()
    Write-Host "  Result: $message" -ForegroundColor Green

    $sysinfo = [TestClass]::GetSystemInfo()
    Write-Host "  System: $sysinfo" -ForegroundColor Green
    Write-Host "  PASSED" -ForegroundColor Green
    Write-Host ""

    # Test 2: Add-Type with System.Runtime.InteropServices (needed for COM)
    Write-Host "[Test 2] InteropServices and COM attributes..." -ForegroundColor Yellow

    $comTestCode = @"
using System;
using System.Runtime.InteropServices;

public class ComTestClass
{
    // Test that we can define COM interface attributes
    [ComImport]
    [Guid("00000000-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IUnknown
    {
        IntPtr QueryInterface(ref Guid riid, out IntPtr ppvObject);
        uint AddRef();
        uint Release();
    }

    // Test struct with LayoutKind
    [StructLayout(LayoutKind.Sequential)]
    public struct TestStruct
    {
        public Guid Id;
        public uint Value;
    }

    public static string TestInterop()
    {
        var ts = new TestStruct();
        ts.Id = Guid.NewGuid();
        ts.Value = 42;
        return string.Format("Struct created - GUID: {0}, Value: {1}", ts.Id, ts.Value);
    }
}
"@

    Add-Type -TypeDefinition $comTestCode -Language CSharp

    $interopResult = [ComTestClass]::TestInterop()
    Write-Host "  Result: $interopResult" -ForegroundColor Green
    Write-Host "  PASSED" -ForegroundColor Green
    Write-Host ""

    # Test 3: COM object instantiation (if this works, IMMDeviceEnumerator will work)
    Write-Host "[Test 3] COM object creation..." -ForegroundColor Yellow

    # Try to create a simple COM object that exists on all Windows systems
    $shell = New-Object -ComObject Shell.Application
    if ($shell) {
        Write-Host "  Shell.Application COM object created successfully" -ForegroundColor Green
        Write-Host "  PASSED" -ForegroundColor Green
    }
    Write-Host ""

    # All tests passed
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "SUCCESS: All tests passed!" -ForegroundColor Green
    Write-Host "Add-Type with inline C# works on this system." -ForegroundColor Green
    Write-Host "The Bluetooth audio script should work." -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green

    exit 0
}
catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "FAILED: Add-Type test failed" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "This means the Bluetooth script may not work" -ForegroundColor Red
    Write-Host "on this locked-down system." -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red

    exit 1
}
