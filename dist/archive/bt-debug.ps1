<#
.SYNOPSIS
    Debug version - verbose output to diagnose issues
#>

param(
    [Parameter(Position=0)]
    [string]$Action = "list",

    [Parameter(Position=1)]
    [string]$DeviceName = ""
)

Write-Host "=== Bluetooth Audio Debug Script ===" -ForegroundColor Cyan
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
Write-Host "Action: $Action" -ForegroundColor Gray
Write-Host "DeviceName: $DeviceName" -ForegroundColor Gray
Write-Host ""

# Step 1: Test basic Add-Type
Write-Host "[Step 1] Testing basic Add-Type..." -ForegroundColor Yellow
try {
    Add-Type -TypeDefinition "public class TestBasic { public static string Test() { return \"OK\"; } }"
    $result = [TestBasic]::Test()
    Write-Host "  Basic Add-Type: $result" -ForegroundColor Green
} catch {
    Write-Host "  FAILED: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Add-Type is blocked on this system." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Step 2: Test COM creation
Write-Host "[Step 2] Testing COM object creation..." -ForegroundColor Yellow
try {
    $shell = New-Object -ComObject Shell.Application
    Write-Host "  COM object creation: OK" -ForegroundColor Green
} catch {
    Write-Host "  FAILED: $_" -ForegroundColor Red
}

# Step 3: Compile the audio code
Write-Host "[Step 3] Compiling Bluetooth audio module..." -ForegroundColor Yellow

$csharpCode = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace BluetoothAudioDebug
{
    public enum EDataFlow : uint { eRender = 0, eCapture = 1, eAll = 2 }

    [Flags]
    public enum DEVICE_STATE : uint
    {
        ACTIVE = 0x00000001,
        DISABLED = 0x00000002,
        NOTPRESENT = 0x00000004,
        UNPLUGGED = 0x00000008,
        MASK_ALL = 0x0000000F
    }

    public enum STGM : uint { READ = 0x00000000 }

    public enum KSPROPERTY_BTAUDIO : uint { ONESHOT_RECONNECT = 0, ONESHOT_DISCONNECT = 1 }

    public enum KsPropertyKind : uint { GET = 0x00000001 }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROPERTYKEY
    {
        public Guid fmtid;
        public uint pid;
        public PROPERTYKEY(Guid guid, uint id) { fmtid = guid; pid = id; }
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct PROPVARIANT
    {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr pwszVal;
        public string GetString()
        {
            if (vt == 31) return Marshal.PtrToStringUni(pwszVal);
            return null;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct KsProperty
    {
        public Guid Set;
        public KSPROPERTY_BTAUDIO Id;
        public KsPropertyKind Flags;
        public KsProperty(Guid set, KSPROPERTY_BTAUDIO id, KsPropertyKind flags)
        { Set = set; Id = id; Flags = flags; }
    }

    [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceCollection
    {
        int GetCount(out uint count);
        int Item(uint index, out IMMDevice device);
    }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice
    {
        int Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, uint clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        int OpenPropertyStore(STGM stgmAccess, out IPropertyStore properties);
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetState(out DEVICE_STATE state);
    }

    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceEnumerator
    {
        int EnumAudioEndpoints(EDataFlow dataFlow, DEVICE_STATE stateMask, out IMMDeviceCollection devices);
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, uint role, out IMMDevice device);
        int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string id, out IMMDevice device);
    }

    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore
    {
        int GetCount(out uint count);
        int GetAt(uint index, out PROPERTYKEY key);
        int GetValue(ref PROPERTYKEY key, out PROPVARIANT value);
    }

    [ComImport, Guid("2A07407E-6497-4A18-9787-32F79BD0D98F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeviceTopology
    {
        int GetConnectorCount(out uint count);
        int GetConnector(uint index, out IConnector connector);
        int GetSubunitCount(out uint count);
        int GetSubunit(uint index, [MarshalAs(UnmanagedType.IUnknown)] out object subunit);
        int GetPartById(uint id, out IPart part);
        int GetDeviceId([MarshalAs(UnmanagedType.LPWStr)] out string deviceId);
        int GetSignalPath([MarshalAs(UnmanagedType.IUnknown)] object partFrom, [MarshalAs(UnmanagedType.IUnknown)] object partTo, bool rejectMixedPaths, [MarshalAs(UnmanagedType.IUnknown)] out object parts);
    }

    [ComImport, Guid("9C2C4058-23F5-41DE-877A-DF3AF236A09E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IConnector
    {
        int GetType(out uint type);
        int GetDataFlow(out uint flow);
        int ConnectTo([MarshalAs(UnmanagedType.IUnknown)] object connector);
        int Disconnect();
        int IsConnected(out bool connected);
        int GetConnectedTo(out IConnector connector);
        int GetConnectorIdConnectedTo([MarshalAs(UnmanagedType.LPWStr)] out string connectorId);
        int GetDeviceIdConnectedTo([MarshalAs(UnmanagedType.LPWStr)] out string deviceId);
    }

    [ComImport, Guid("AE2DE0E4-5BCA-4F2D-AA46-5D13F8FDB3A9"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPart
    {
        int GetName([MarshalAs(UnmanagedType.LPWStr)] out string name);
        int GetLocalId(out uint id);
        int GetGlobalId([MarshalAs(UnmanagedType.LPWStr)] out string globalId);
        int GetPartType(out uint partType);
        int GetSubType(out Guid subType);
        int GetControlInterfaceCount(out uint count);
        int GetControlInterface(uint index, [MarshalAs(UnmanagedType.IUnknown)] out object controlInterface);
        int EnumPartsIncoming([MarshalAs(UnmanagedType.IUnknown)] out object parts);
        int EnumPartsOutgoing([MarshalAs(UnmanagedType.IUnknown)] out object parts);
        int GetTopologyObject(out IDeviceTopology topology);
        int Activate(uint clsContext, [MarshalAs(UnmanagedType.LPStruct)] Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object activated);
    }

    [ComImport, Guid("28F54685-06FD-11D2-B27A-00A0C9223196"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IKsControl
    {
        int KsProperty(ref KsProperty property, int propertyLength, IntPtr propertyData, int dataLength, ref int bytesReturned);
        int KsMethod(ref KsProperty method, int methodLength, IntPtr methodData, int dataLength, ref int bytesReturned);
        int KsEvent(ref KsProperty evt, int eventLength, IntPtr eventData, int dataLength, ref int bytesReturned);
    }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    public class MMDeviceEnumerator { }

    public class BtAudioDevice
    {
        public string Name;
        public string Id;
        public bool IsActive;
        public IKsControl KsControl;
    }

    public static class BtAudioManager
    {
        private static readonly Guid KSPROPSETID_BtAudio = new Guid("7fa06c40-b8f6-4c7e-8556-e8c33a12e54d");
        private static readonly Guid IID_IDeviceTopology = new Guid("2A07407E-6497-4A18-9787-32F79BD0D98F");
        private static readonly Guid IID_IKsControl = new Guid("28F54685-06FD-11D2-B27A-00A0C9223196");
        private static readonly PROPERTYKEY PKEY_Device_FriendlyName = new PROPERTYKEY(
            new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), 14);

        public static string GetDebugInfo()
        {
            var sb = new System.Text.StringBuilder();
            try
            {
                sb.AppendLine("Creating MMDeviceEnumerator...");
                var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
                sb.AppendLine("  OK");

                sb.AppendLine("Enumerating audio endpoints...");
                IMMDeviceCollection collection;
                int hr = enumerator.EnumAudioEndpoints(EDataFlow.eAll, DEVICE_STATE.MASK_ALL, out collection);
                sb.AppendLine("  EnumAudioEndpoints returned: 0x" + hr.ToString("X8"));

                uint count;
                collection.GetCount(out count);
                sb.AppendLine("  Found " + count + " audio endpoints");

                int btCount = 0;
                for (uint i = 0; i < count; i++)
                {
                    try
                    {
                        IMMDevice device;
                        collection.Item(i, out device);

                        string name = GetDeviceName(device);
                        string deviceId;
                        device.GetId(out deviceId);
                        DEVICE_STATE state;
                        device.GetState(out state);

                        sb.AppendLine("");
                        sb.AppendLine("  [" + i + "] " + name);
                        sb.AppendLine("      State: " + state);
                        sb.AppendLine("      ID: " + (deviceId != null && deviceId.Length > 50 ? deviceId.Substring(0, 50) + "..." : deviceId));

                        // Try to get topology
                        try
                        {
                            object topologyObj;
                            hr = device.Activate(IID_IDeviceTopology, 1, IntPtr.Zero, out topologyObj);
                            if (hr >= 0 && topologyObj != null)
                            {
                                var topology = (IDeviceTopology)topologyObj;
                                uint connCount;
                                topology.GetConnectorCount(out connCount);
                                sb.AppendLine("      Connectors: " + connCount);

                                for (uint c = 0; c < connCount; c++)
                                {
                                    try
                                    {
                                        IConnector connector;
                                        topology.GetConnector(c, out connector);

                                        bool isConnected;
                                        connector.IsConnected(out isConnected);

                                        if (isConnected)
                                        {
                                            IConnector connectedTo;
                                            connector.GetConnectedTo(out connectedTo);

                                            string connectedDeviceId;
                                            connectedTo.GetDeviceIdConnectedTo(out connectedDeviceId);

                                            sb.AppendLine("      Connected to: " + (connectedDeviceId != null && connectedDeviceId.Length > 40 ? connectedDeviceId.Substring(0, 40) + "..." : connectedDeviceId));

                                            if (connectedDeviceId != null && connectedDeviceId.Contains("bth"))
                                            {
                                                sb.AppendLine("      *** BLUETOOTH DEVICE ***");
                                                btCount++;
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        sb.AppendLine("      Connector " + c + " error: " + ex.Message);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            sb.AppendLine("      Topology error: " + ex.Message);
                        }
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("  [" + i + "] Error: " + ex.Message);
                    }
                }

                sb.AppendLine("");
                sb.AppendLine("=== Summary ===");
                sb.AppendLine("Total audio endpoints: " + count);
                sb.AppendLine("Bluetooth devices found: " + btCount);
            }
            catch (Exception ex)
            {
                sb.AppendLine("FATAL ERROR: " + ex.GetType().Name + ": " + ex.Message);
                sb.AppendLine(ex.StackTrace);
            }
            return sb.ToString();
        }

        private static string GetDeviceName(IMMDevice device)
        {
            try
            {
                IPropertyStore props;
                device.OpenPropertyStore(STGM.READ, out props);
                PROPVARIANT value;
                var key = PKEY_Device_FriendlyName;
                props.GetValue(ref key, out value);
                return value.GetString() ?? "<no name>";
            }
            catch
            {
                return "<error getting name>";
            }
        }
    }
}
"@

try {
    Add-Type -TypeDefinition $csharpCode -Language CSharp
    Write-Host "  Compilation: OK" -ForegroundColor Green
} catch {
    Write-Host "  FAILED to compile:" -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Full error:" -ForegroundColor Red
    Write-Host $_.Exception.ToString() -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Step 4: Run diagnostics
Write-Host "[Step 4] Enumerating audio devices..." -ForegroundColor Yellow
Write-Host ""

try {
    $debugInfo = [BluetoothAudioDebug.BtAudioManager]::GetDebugInfo()
    Write-Host $debugInfo
} catch {
    Write-Host "FAILED to enumerate devices:" -ForegroundColor Red
    Write-Host $_.Exception.ToString() -ForegroundColor Red
}

Write-Host ""
Write-Host "=== Debug Complete ===" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"
