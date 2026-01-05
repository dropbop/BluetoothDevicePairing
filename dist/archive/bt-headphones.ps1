<#
.SYNOPSIS
    Connect or disconnect Bluetooth audio devices.

.DESCRIPTION
    A self-contained script to manage Bluetooth audio device connections.
    Uses Windows Core Audio APIs via inline C# - no external dependencies.

.PARAMETER Action
    The action to perform: list, connect, or disconnect

.PARAMETER DeviceName
    The name (or partial name) of the Bluetooth audio device

.EXAMPLE
    .\bt-headphones.ps1 list
    Lists all Bluetooth audio devices

.EXAMPLE
    .\bt-headphones.ps1 connect "WH-1000XM5"
    Connects to the specified device

.EXAMPLE
    .\bt-headphones.ps1 disconnect "WH-1000XM5"
    Disconnects the specified device
#>

param(
    [Parameter(Position=0)]
    [ValidateSet("list", "connect", "disconnect", "help")]
    [string]$Action = "help",

    [Parameter(Position=1)]
    [string]$DeviceName = ""
)

$ErrorActionPreference = "Stop"

# ============================================================================
# Inline C# Code - Core Audio COM Interfaces for Bluetooth Audio Control
# ============================================================================

$csharpCode = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace BluetoothAudio
{
    // ========================================================================
    // Constants and Enums
    // ========================================================================

    public enum EDataFlow : uint
    {
        eRender = 0,
        eCapture = 1,
        eAll = 2
    }

    [Flags]
    public enum DEVICE_STATE : uint
    {
        ACTIVE = 0x00000001,
        DISABLED = 0x00000002,
        NOTPRESENT = 0x00000004,
        UNPLUGGED = 0x00000008,
        MASK_ALL = 0x0000000F
    }

    public enum STGM : uint
    {
        READ = 0x00000000
    }

    public enum KSPROPERTY_BTAUDIO : uint
    {
        ONESHOT_RECONNECT = 0,
        ONESHOT_DISCONNECT = 1
    }

    public enum KsPropertyKind : uint
    {
        GET = 0x00000001,
        SET = 0x00000002,
        TOPOLOGY = 0x10000000
    }

    // ========================================================================
    // Structs
    // ========================================================================

    [StructLayout(LayoutKind.Sequential)]
    public struct PROPERTYKEY
    {
        public Guid fmtid;
        public uint pid;

        public PROPERTYKEY(Guid guid, uint id)
        {
            fmtid = guid;
            pid = id;
        }
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct PROPVARIANT
    {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr pwszVal;

        public string GetString()
        {
            if (vt == 31) // VT_LPWSTR
                return Marshal.PtrToStringUni(pwszVal);
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
        {
            Set = set;
            Id = id;
            Flags = flags;
        }
    }

    // ========================================================================
    // COM Interface Definitions
    // ========================================================================

    [ComImport]
    [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceCollection
    {
        int GetCount(out uint count);
        int Item(uint index, out IMMDevice device);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice
    {
        int Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, uint clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        int OpenPropertyStore(STGM stgmAccess, out IPropertyStore properties);
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetState(out DEVICE_STATE state);
    }

    [ComImport]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceEnumerator
    {
        int EnumAudioEndpoints(EDataFlow dataFlow, DEVICE_STATE stateMask, out IMMDeviceCollection devices);
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, uint role, out IMMDevice device);
        int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string id, out IMMDevice device);
    }

    [ComImport]
    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore
    {
        int GetCount(out uint count);
        int GetAt(uint index, out PROPERTYKEY key);
        int GetValue(ref PROPERTYKEY key, out PROPVARIANT value);
    }

    [ComImport]
    [Guid("2A07407E-6497-4A18-9787-32F79BD0D98F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
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

    [ComImport]
    [Guid("9C2C4058-23F5-41DE-877A-DF3AF236A09E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IConnector
    {
        int QueryInterface(ref Guid riid, out IntPtr ppvObject);
        int GetType(out uint type);
        int GetDataFlow(out uint flow);
        int ConnectTo([MarshalAs(UnmanagedType.IUnknown)] object connector);
        int Disconnect();
        int IsConnected(out bool connected);
        int GetConnectedTo(out IConnector connector);
        int GetConnectorIdConnectedTo([MarshalAs(UnmanagedType.LPWStr)] out string connectorId);
        int GetDeviceIdConnectedTo([MarshalAs(UnmanagedType.LPWStr)] out string deviceId);
    }

    [ComImport]
    [Guid("AE2DE0E4-5BCA-4F2D-AA46-5D13F8FDB3A9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
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

    [ComImport]
    [Guid("28F54685-06FD-11D2-B27A-00A0C9223196")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IKsControl
    {
        int KsProperty(ref KsProperty property, int propertyLength, IntPtr propertyData, int dataLength, ref int bytesReturned);
        int KsMethod(ref KsProperty method, int methodLength, IntPtr methodData, int dataLength, ref int bytesReturned);
        int KsEvent(ref KsProperty evt, int eventLength, IntPtr eventData, int dataLength, ref int bytesReturned);
    }

    // ========================================================================
    // MMDeviceEnumerator COM Class
    // ========================================================================

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    public class MMDeviceEnumerator { }

    // ========================================================================
    // Bluetooth Audio Device Info
    // ========================================================================

    public class BluetoothAudioDevice
    {
        public string Name { get; set; }
        public string Id { get; set; }
        public bool IsActive { get; set; }
        public IKsControl KsControl { get; set; }

        public override string ToString()
        {
            string status = IsActive ? "Connected" : "Disconnected";
            return string.Format("{0} [{1}]", Name, status);
        }
    }

    // ========================================================================
    // Main Manager Class
    // ========================================================================

    public static class BluetoothAudioManager
    {
        private static readonly Guid KSPROPSETID_BtAudio = new Guid("7fa06c40-b8f6-4c7e-8556-e8c33a12e54d");
        private static readonly Guid IID_IDeviceTopology = new Guid("2A07407E-6497-4A18-9787-32F79BD0D98F");
        private static readonly Guid IID_IKsControl = new Guid("28F54685-06FD-11D2-B27A-00A0C9223196");
        private static readonly PROPERTYKEY PKEY_Device_FriendlyName = new PROPERTYKEY(
            new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), 14);

        public static List<BluetoothAudioDevice> GetBluetoothAudioDevices()
        {
            var devices = new List<BluetoothAudioDevice>();

            try
            {
                var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
                IMMDeviceCollection collection;
                enumerator.EnumAudioEndpoints(EDataFlow.eAll, DEVICE_STATE.MASK_ALL, out collection);

                uint count;
                collection.GetCount(out count);

                for (uint i = 0; i < count; i++)
                {
                    try
                    {
                        IMMDevice device;
                        collection.Item(i, out device);

                        var btDevice = TryGetBluetoothDevice(device, enumerator);
                        if (btDevice != null)
                        {
                            devices.Add(btDevice);
                        }
                    }
                    catch
                    {
                        // Skip devices that can't be accessed
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to enumerate audio devices: " + ex.Message);
            }

            return devices;
        }

        private static BluetoothAudioDevice TryGetBluetoothDevice(IMMDevice device, IMMDeviceEnumerator enumerator)
        {
            try
            {
                // Get device topology
                object topologyObj;
                device.Activate(IID_IDeviceTopology, 1, IntPtr.Zero, out topologyObj);
                var topology = (IDeviceTopology)topologyObj;

                uint connectorCount;
                topology.GetConnectorCount(out connectorCount);

                for (uint c = 0; c < connectorCount; c++)
                {
                    try
                    {
                        IConnector connector;
                        topology.GetConnector(c, out connector);

                        bool isConnected;
                        connector.IsConnected(out isConnected);
                        if (!isConnected) continue;

                        IConnector connectedTo;
                        connector.GetConnectedTo(out connectedTo);

                        // Get the connected device ID
                        string connectedDeviceId;
                        connectedTo.GetDeviceIdConnectedTo(out connectedDeviceId);

                        // Check if it's a Bluetooth device
                        if (connectedDeviceId != null && connectedDeviceId.StartsWith(@"{2}.\\?\bth", StringComparison.OrdinalIgnoreCase))
                        {
                            // Get the connected Bluetooth device
                            IMMDevice btDevice;
                            enumerator.GetDevice(connectedDeviceId, out btDevice);

                            // Get IKsControl from the Bluetooth device
                            object ksControlObj;
                            btDevice.Activate(IID_IKsControl, 1, IntPtr.Zero, out ksControlObj);
                            var ksControl = (IKsControl)ksControlObj;

                            // Get device name and state from original endpoint
                            string name = GetDeviceName(device);
                            DEVICE_STATE state;
                            device.GetState(out state);

                            return new BluetoothAudioDevice
                            {
                                Name = name,
                                Id = connectedDeviceId,
                                IsActive = (state & DEVICE_STATE.ACTIVE) != 0,
                                KsControl = ksControl
                            };
                        }
                    }
                    catch
                    {
                        // Skip this connector
                    }
                }
            }
            catch
            {
                // Not a Bluetooth device or can't get topology
            }

            return null;
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

                return value.GetString() ?? "Unknown Device";
            }
            catch
            {
                return "Unknown Device";
            }
        }

        public static string ListDevices()
        {
            var devices = GetBluetoothAudioDevices();

            if (devices.Count == 0)
            {
                return "No Bluetooth audio devices found.\n\nMake sure your device is:\n  - Paired with this computer\n  - Turned on and in range";
            }

            var result = new System.Text.StringBuilder();
            result.AppendLine("Bluetooth Audio Devices:");
            result.AppendLine(new string('-', 50));

            foreach (var device in devices)
            {
                result.AppendLine(device.ToString());
            }

            return result.ToString();
        }

        public static string Connect(string deviceName)
        {
            return SendKsProperty(deviceName, KSPROPERTY_BTAUDIO.ONESHOT_RECONNECT, "connect");
        }

        public static string Disconnect(string deviceName)
        {
            return SendKsProperty(deviceName, KSPROPERTY_BTAUDIO.ONESHOT_DISCONNECT, "disconnect");
        }

        private static string SendKsProperty(string deviceName, KSPROPERTY_BTAUDIO property, string action)
        {
            if (string.IsNullOrWhiteSpace(deviceName))
            {
                return "Error: Device name is required.\n\nUsage: bt-headphones.ps1 " + action + " \"DeviceName\"\n\nRun 'bt-headphones.ps1 list' to see available devices.";
            }

            var devices = GetBluetoothAudioDevices();

            if (devices.Count == 0)
            {
                return "No Bluetooth audio devices found.\n\nMake sure your device is paired and turned on.";
            }

            // Find matching devices (case-insensitive partial match)
            var matches = devices.FindAll(d =>
                d.Name != null && d.Name.IndexOf(deviceName, StringComparison.OrdinalIgnoreCase) >= 0);

            if (matches.Count == 0)
            {
                var result = new System.Text.StringBuilder();
                result.AppendLine("No device found matching: " + deviceName);
                result.AppendLine();
                result.AppendLine("Available devices:");
                foreach (var d in devices)
                {
                    result.AppendLine("  - " + d.Name);
                }
                return result.ToString();
            }

            if (matches.Count > 1)
            {
                var result = new System.Text.StringBuilder();
                result.AppendLine("Multiple devices match '" + deviceName + "':");
                foreach (var d in matches)
                {
                    result.AppendLine("  - " + d.Name);
                }
                result.AppendLine();
                result.AppendLine("Please use a more specific name.");
                return result.ToString();
            }

            var device = matches[0];

            try
            {
                var ksProperty = new KsProperty(KSPROPSETID_BtAudio, property, KsPropertyKind.GET);
                int bytesReturned = 0;
                int hr = device.KsControl.KsProperty(ref ksProperty, Marshal.SizeOf(ksProperty), IntPtr.Zero, 0, ref bytesReturned);

                if (hr >= 0)
                {
                    string actionPast = property == KSPROPERTY_BTAUDIO.ONESHOT_RECONNECT ? "Connection" : "Disconnection";
                    return actionPast + " request sent to: " + device.Name + "\n\nNote: It may take a few seconds for the device to respond.";
                }
                else
                {
                    return "Failed to " + action + " device: " + device.Name + "\nError code: 0x" + hr.ToString("X8");
                }
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }
    }
}
"@

# ============================================================================
# Compile the C# code
# ============================================================================

try {
    Add-Type -TypeDefinition $csharpCode -Language CSharp
}
catch {
    Write-Host "Failed to compile Bluetooth audio module:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

# ============================================================================
# Main Script Logic
# ============================================================================

function Show-Help {
    Write-Host ""
    Write-Host "Bluetooth Audio Device Manager" -ForegroundColor Cyan
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  .\bt-headphones.ps1 list                    List Bluetooth audio devices"
    Write-Host "  .\bt-headphones.ps1 connect <name>          Connect to a device"
    Write-Host "  .\bt-headphones.ps1 disconnect <name>       Disconnect a device"
    Write-Host "  .\bt-headphones.ps1 help                    Show this help"
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host '  .\bt-headphones.ps1 list'
    Write-Host '  .\bt-headphones.ps1 connect "WH-1000XM5"'
    Write-Host '  .\bt-headphones.ps1 disconnect "AirPods"'
    Write-Host ""
    Write-Host "Notes:" -ForegroundColor Yellow
    Write-Host "  - Device names support partial matching (case-insensitive)"
    Write-Host "  - Devices must be paired first via Windows Settings"
    Write-Host ""
}

switch ($Action.ToLower()) {
    "list" {
        $result = [BluetoothAudio.BluetoothAudioManager]::ListDevices()
        Write-Host $result
    }
    "connect" {
        $result = [BluetoothAudio.BluetoothAudioManager]::Connect($DeviceName)
        Write-Host $result
    }
    "disconnect" {
        $result = [BluetoothAudio.BluetoothAudioManager]::Disconnect($DeviceName)
        Write-Host $result
    }
    "help" {
        Show-Help
    }
    default {
        Show-Help
    }
}
