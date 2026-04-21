#Requires -Version 5.1
# ==================================================================
# Audio Manager v26.04.21 - by PeterYama
# iwr -useb https://raw.githubusercontent.com/PeterYama/
#     audio-manager/master/AudioManager.ps1 | iex
# https://github.com/PeterYama/audio-manager
# ==================================================================
$script:AMVersion = "26.04.21"

# --- Elevation check ---

$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin   = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    if ($PSCommandPath) {
        $target = $PSCommandPath
    } else {
        # Script was IEX'd - save to temp and relaunch elevated
        $target = "$env:TEMP\AudioManager_elevated.ps1"
        try {
            $scriptContent = (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/PeterYama/audio-manager/master/AudioManager.ps1" -UseBasicParsing)
            Set-Content -Path $target -Value $scriptContent -Encoding UTF8
        } catch {
            Write-Host "Could not auto-download for elevation. Please run PowerShell as Administrator." -ForegroundColor Red
            Start-Sleep -Seconds 4
            exit 1
        }
    }
    $relaunchArgs = "-ExecutionPolicy Bypass -File `"$target`""
    Start-Process powershell.exe -Verb RunAs -ArgumentList $relaunchArgs
    exit
}

# --- Assembly loading ---

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# --- Logging ---

$logDir = "$env:LOCALAPPDATA\AudioManager\logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
Start-Transcript -Path "$logDir\AudioManager_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" -Append -ErrorAction SilentlyContinue

# --- Profiles directory ---

$profilesDir = "$env:APPDATA\AudioManager"
if (-not (Test-Path $profilesDir)) { New-Item -ItemType Directory -Path $profilesDir -Force | Out-Null }

# --- Shared synchronized hashtable ---

$sync = [hashtable]::Synchronized(@{
    Form              = $null
    CurrentTab        = "Devices"
    RenderDevices     = @()
    CaptureDevices    = @()
    AudioSessions     = @()
    Profiles          = @()
    SelectedOutputId  = $null
    SelectedInputId   = $null
    IsRefreshing      = $false
    ProfilesPath      = "$env:APPDATA\AudioManager\profiles.json"
    Version           = $script:AMVersion
    NullGuid          = [guid]::Empty
})

# RunspacePool is created in main.ps1 after all functions are defined.

$script:CoreAudioCSharp = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace AudioManager
{
    // -----------------------------------------------------------------------
    // Enumerations
    // -----------------------------------------------------------------------

    public enum EDataFlow   { eRender = 0, eCapture = 1, eAll = 2 }
    public enum ERole       { eConsole = 0, eMultimedia = 1, eCommunications = 2 }
    public enum AudioSessionState { Inactive = 0, Active = 1, Expired = 2 }

    [Flags]
    public enum DeviceState : uint
    {
        Active     = 0x00000001,
        Disabled   = 0x00000002,
        NotPresent = 0x00000004,
        Unplugged  = 0x00000008,
        All        = 0x0000000F
    }

    // -----------------------------------------------------------------------
    // Structs
    // -----------------------------------------------------------------------

    [StructLayout(LayoutKind.Sequential)]
    internal struct WAVEFORMATEX
    {
        public ushort wFormatTag;
        public ushort nChannels;
        public uint   nSamplesPerSec;
        public uint   nAvgBytesPerSec;
        public ushort nBlockAlign;
        public ushort wBitsPerSample;
        public ushort cbSize;
    }

    [StructLayout(LayoutKind.Explicit, Pack = 1)]
    internal struct WAVEFORMATEXTENSIBLE
    {
        [FieldOffset(0)]  public WAVEFORMATEX Format;
        [FieldOffset(18)] public ushort       wValidBitsPerSample;
        [FieldOffset(20)] public uint         dwChannelMask;
        [FieldOffset(24)] public Guid         SubFormat;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PropertyKey
    {
        public Guid fmtid;
        public uint pid;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PropVariant
    {
        public ushort vt;
        public ushort reserved1;
        public ushort reserved2;
        public ushort reserved3;
        public IntPtr data;
        public IntPtr data2;

        public string GetStringValue()
        {
            if (vt == 31) // VT_LPWSTR
                return Marshal.PtrToStringUni(data);
            return string.Empty;
        }
    }

    // -----------------------------------------------------------------------
    // COM Interfaces  (internal - never exposed directly to PowerShell)
    // -----------------------------------------------------------------------

    // IMMDeviceEnumerator -- EnumAudioEndpoints returns IntPtr because the
    // IMMDeviceCollection COM object does not respond to QI for its own IID
    // when accessed from managed code; we use raw vtable dispatch instead.
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IMMDeviceEnumerator
    {
        [PreserveSig]
        int EnumAudioEndpoints([In] EDataFlow flow, [In] DeviceState mask,
                               [Out] out IntPtr ppDevices);
        [PreserveSig]
        int GetDefaultAudioEndpoint([In] EDataFlow flow, [In] ERole role,
                                    [Out] out IMMDevice endpoint);
        [PreserveSig]
        int GetDevice([In, MarshalAs(UnmanagedType.LPWStr)] string id,
                      [Out] out IMMDevice device);
        [PreserveSig] int RegisterEndpointNotificationCallback(IntPtr client);
        [PreserveSig] int UnregisterEndpointNotificationCallback(IntPtr client);
    }

    // IMMDeviceCollection is not used as a managed interface type -- QI for
    // its IID fails at runtime.  All collection access goes through the raw
    // vtable delegates defined in AudioManagerHelper.
    // (keeping the declaration here for documentation only -- it is not used)
    [Guid("0BD7A1BE-7A1A-44DB-8397-BE5155E7F6E1")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IMMDeviceCollection
    {
        int GetCount([Out] out uint count);
        int Item([In] uint index, [Out] out IMMDevice device);
    }

    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IMMDevice
    {
        int Activate([In] ref Guid iid, [In] int clsCtx,
                     [In] IntPtr pActivationParams,
                     [Out, MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        int OpenPropertyStore([In] int stgmAccess, [Out] out IPropertyStore props);
        int GetId([Out, MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetState([Out] out DeviceState state);
    }

    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IPropertyStore
    {
        int GetCount([Out] out int count);
        int GetAt([In] int index, [Out] out PropertyKey key);
        int GetValue([In] ref PropertyKey key, [Out] out PropVariant value);
        int SetValue([In] ref PropertyKey key, [In] ref PropVariant value);
        int Commit();
    }

    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IAudioEndpointVolume
    {
        int RegisterControlChangeNotify(IntPtr cb);
        int UnregisterControlChangeNotify(IntPtr cb);
        int GetChannelCount([Out] out int count);
        int SetMasterVolumeLevel([In] float level, [In] ref Guid ctx);
        int SetMasterVolumeLevelScalar([In] float level, [In] ref Guid ctx);
        int GetMasterVolumeLevel([Out] out float level);
        int GetMasterVolumeLevelScalar([Out] out float level);
        int SetChannelVolumeLevel([In] uint ch, [In] float level, [In] ref Guid ctx);
        int SetChannelVolumeLevelScalar([In] uint ch, [In] float level, [In] ref Guid ctx);
        int GetChannelVolumeLevel([In] uint ch, [Out] out float level);
        int GetChannelVolumeLevelScalar([In] uint ch, [Out] out float level);
        int SetMute([In, MarshalAs(UnmanagedType.Bool)] bool muted, [In] ref Guid ctx);
        int GetMute([Out, MarshalAs(UnmanagedType.Bool)] out bool muted);
        int GetVolumeStepInfo([Out] out uint step, [Out] out uint count);
        int VolumeStepUp([In] ref Guid ctx);
        int VolumeStepDown([In] ref Guid ctx);
        int QueryHardwareSupport([Out] out uint mask);
        int GetVolumeRange([Out] out float min, [Out] out float max, [Out] out float inc);
    }

    [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IAudioSessionManager2
    {
        int GetAudioSessionControl([In] ref Guid sessionGuid, [In] uint flags,
                                   [Out] out IAudioSessionControl session);
        int GetSimpleAudioVolume([In] ref Guid sessionGuid, [In] uint flags,
                                 [Out] out ISimpleAudioVolume vol);
        int GetSessionEnumerator([Out] out IAudioSessionEnumerator sessionEnum);
        int RegisterSessionNotification(IntPtr notify);
        int UnregisterSessionNotification(IntPtr notify);
        int RegisterDuckNotification([MarshalAs(UnmanagedType.LPWStr)] string id, IntPtr notify);
        int UnregisterDuckNotification(IntPtr notify);
    }

    [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IAudioSessionEnumerator
    {
        int GetCount([Out] out int count);
        int GetSession([In] int index, [Out] out IAudioSessionControl session);
    }

    [Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IAudioSessionControl
    {
        int GetState([Out] out AudioSessionState state);
        int GetDisplayName([Out, MarshalAs(UnmanagedType.LPWStr)] out string name);
        int SetDisplayName([In, MarshalAs(UnmanagedType.LPWStr)] string val, [In] ref Guid ctx);
        int GetIconPath([Out, MarshalAs(UnmanagedType.LPWStr)] out string path);
        int SetIconPath([In, MarshalAs(UnmanagedType.LPWStr)] string val, [In] ref Guid ctx);
        int GetGroupingParam([Out] out Guid param);
        int SetGroupingParam([In] ref Guid param, [In] ref Guid ctx);
        int RegisterAudioSessionNotification(IntPtr notify);
        int UnregisterAudioSessionNotification(IntPtr notify);
    }

    [Guid("BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IAudioSessionControl2
    {
        int GetState([Out] out AudioSessionState state);
        int GetDisplayName([Out, MarshalAs(UnmanagedType.LPWStr)] out string name);
        int SetDisplayName([In, MarshalAs(UnmanagedType.LPWStr)] string val, [In] ref Guid ctx);
        int GetIconPath([Out, MarshalAs(UnmanagedType.LPWStr)] out string path);
        int SetIconPath([In, MarshalAs(UnmanagedType.LPWStr)] string val, [In] ref Guid ctx);
        int GetGroupingParam([Out] out Guid param);
        int SetGroupingParam([In] ref Guid param, [In] ref Guid ctx);
        int RegisterAudioSessionNotification(IntPtr notify);
        int UnregisterAudioSessionNotification(IntPtr notify);
        int GetSessionIdentifier([Out, MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetSessionInstanceIdentifier([Out, MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetProcessId([Out] out uint pid);
        int IsSystemSoundsSession();
        int SetDuckingPreference([In, MarshalAs(UnmanagedType.Bool)] bool optOut);
    }

    [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface ISimpleAudioVolume
    {
        int SetMasterVolume([In] float level, [In] ref Guid ctx);
        int GetMasterVolume([Out] out float level);
        int SetMute([In, MarshalAs(UnmanagedType.Bool)] bool muted, [In] ref Guid ctx);
        int GetMute([Out, MarshalAs(UnmanagedType.Bool)] out bool muted);
    }

    [Guid("F8679F50-850A-41CF-9C72-430F290290C8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IPolicyConfig
    {
        int GetMixFormat([MarshalAs(UnmanagedType.LPWStr)] string dev, [Out] out IntPtr fmt);
        int GetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string dev,
                            [In, MarshalAs(UnmanagedType.Bool)] bool bDefault,
                            [Out] out IntPtr fmt);
        int ResetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string dev);
        int SetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string dev,
                            [In] IntPtr endpointFmt, IntPtr mixFmt);
        int GetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string dev,
                                [In, MarshalAs(UnmanagedType.Bool)] bool bDefault,
                                [Out] out long defPeriod, [Out] out long minPeriod);
        int SetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string dev,
                                [In] ref long period);
        int GetShareMode([MarshalAs(UnmanagedType.LPWStr)] string dev,
                         [Out] out int mode);
        int SetShareMode([MarshalAs(UnmanagedType.LPWStr)] string dev, [In] int mode);
        int GetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string dev,
                             [In, MarshalAs(UnmanagedType.Bool)] bool bFxStore,
                             [In] ref PropertyKey key, [Out] out PropVariant pv);
        int SetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string dev,
                             [In, MarshalAs(UnmanagedType.Bool)] bool bFxStore,
                             [In] ref PropertyKey key, [In] ref PropVariant pv);
        int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string dev, [In] ERole role);
        int SetEndpointVisibility([MarshalAs(UnmanagedType.LPWStr)] string dev,
                                  [In, MarshalAs(UnmanagedType.Bool)] bool visible);
    }

    // -----------------------------------------------------------------------
    // CoClass wrappers  (internal)
    // -----------------------------------------------------------------------

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumeratorComObject { }

    [ComImport, Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
    internal class CPolicyConfigClient { }

    // -----------------------------------------------------------------------
    // Public data classes returned to PowerShell
    // -----------------------------------------------------------------------

    public class AudioDeviceInfo
    {
        public string DeviceId    { get; set; }
        public string Name        { get; set; }
        public float  VolumeScalar { get; set; }
        public bool   IsMuted     { get; set; }
    }

    public class AudioSessionInfo
    {
        public string SessionKey   { get; set; }
        public string Name         { get; set; }
        public int    ProcessId    { get; set; }
        public string PidLabel     { get; set; }
        public int    VolumePercent { get; set; }
        public string VolumeLabel  { get; set; }
        public bool   IsMuted      { get; set; }
        public string State        { get; set; }
        public string Icon         { get; set; }
    }

    public class DeviceFormatInfo
    {
        public string DeviceId   { get; set; }
        public int    SampleRate { get; set; }
        public int    BitDepth   { get; set; }
        public int    Channels   { get; set; }
        public string Label      { get; set; }
    }

    // -----------------------------------------------------------------------
    // Main helper - all COM operations happen here, never in PowerShell
    // -----------------------------------------------------------------------

    public static class AudioManagerHelper
    {
        // GUIDs
        private static readonly Guid IID_IAudioEndpointVolume  =
            new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
        private static readonly Guid IID_IAudioSessionManager2 =
            new Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F");

        // PKEY_Device_FriendlyName
        private static readonly PropertyKey PKEY_FriendlyName = new PropertyKey
        {
            fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"),
            pid   = 14
        };

        // ---- STA thread wrapper -------------------------------------------
        // All IMMDeviceEnumerator / IPolicyConfig COM objects require an STA
        // apartment.  PowerShell runspaces default to MTA, so every public
        // entry-point marshals its work onto a fresh STA thread and returns
        // plain .NET POD objects that cross apartment boundaries freely.

        private static T RunOnSTA<T>(Func<T> work)
        {
            T result = default(T);
            Exception caught = null;
            var thread = new Thread(() =>
            {
                try   { result = work(); }
                catch (Exception ex) { caught = ex; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
            thread.Join();
            if (caught != null) throw new Exception(caught.Message, caught);
            return result;
        }

        private static bool RunOnSTA(Func<bool> work) { return RunOnSTA<bool>(work); }

        // ---- IMMDeviceCollection raw vtable delegates ---------------------
        // QI for IMMDeviceCollection fails at runtime (E_NOINTERFACE) even
        // though the pointer is valid.  We call GetCount / Item directly via
        // the vtable to avoid QI entirely.
        // IMMDeviceCollection vtable layout (inherits IUnknown):
        //   slot 0  QueryInterface
        //   slot 1  AddRef
        //   slot 2  Release
        //   slot 3  GetCount  (IMMDeviceCollection::GetCount)
        //   slot 4  Item      (IMMDeviceCollection::Item)

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int CollGetCountDelegate(IntPtr pThis, out uint count);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int CollItemDelegate(IntPtr pThis, uint index, out IntPtr ppDevice);

        // IMMDevice vtable layout (inherits IUnknown):
        //   slot 0  QueryInterface
        //   slot 1  AddRef
        //   slot 2  Release
        //   slot 3  Activate
        //   slot 4  OpenPropertyStore
        //   slot 5  GetId
        //   slot 6  GetState

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int DevGetIdDelegate(IntPtr pThis,
            [MarshalAs(UnmanagedType.LPWStr)] out string id);

        private static uint CollGetCount(IntPtr col)
        {
            IntPtr vtbl = Marshal.ReadIntPtr(col);
            IntPtr fn   = Marshal.ReadIntPtr(vtbl, 3 * IntPtr.Size);
            uint count  = 0;
            ((CollGetCountDelegate)Marshal.GetDelegateForFunctionPointer(fn, typeof(CollGetCountDelegate)))(col, out count);
            return count;
        }

        private static IntPtr CollItem(IntPtr col, uint i)
        {
            IntPtr vtbl  = Marshal.ReadIntPtr(col);
            IntPtr fn    = Marshal.ReadIntPtr(vtbl, 4 * IntPtr.Size);
            IntPtr devPtr = IntPtr.Zero;
            ((CollItemDelegate)Marshal.GetDelegateForFunctionPointer(fn, typeof(CollItemDelegate)))(col, i, out devPtr);
            return devPtr;
        }

        private static string DevGetIdRaw(IntPtr dev)
        {
            IntPtr vtbl = Marshal.ReadIntPtr(dev);
            IntPtr fn   = Marshal.ReadIntPtr(vtbl, 5 * IntPtr.Size);
            string id   = null;
            ((DevGetIdDelegate)Marshal.GetDelegateForFunctionPointer(fn, typeof(DevGetIdDelegate)))(dev, out id);
            return id;
        }

        // ---- factory helpers -----------------------------------------------

        private static IMMDeviceEnumerator CreateEnumerator()
        {
            return (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
        }

        private static IPolicyConfig CreatePolicyConfig()
        {
            return (IPolicyConfig)new CPolicyConfigClient();
        }

        private static IMMDevice GetDeviceById(string deviceId)
        {
            IMMDevice dev = null;
            CreateEnumerator().GetDevice(deviceId, out dev);
            return dev;
        }

        private static IAudioEndpointVolume GetEndpointVolume(string deviceId)
        {
            var dev  = GetDeviceById(deviceId);
            var iid  = IID_IAudioEndpointVolume;
            object o = null;
            dev.Activate(ref iid, 23, IntPtr.Zero, out o);
            return (IAudioEndpointVolume)o;
        }

        // ---- device enumeration -------------------------------------------

        public static AudioDeviceInfo[] GetRenderDevices()  { return RunOnSTA(() => GetDevices(EDataFlow.eRender));  }
        public static AudioDeviceInfo[] GetCaptureDevices() { return RunOnSTA(() => GetDevices(EDataFlow.eCapture)); }

        private static AudioDeviceInfo[] GetDevices(EDataFlow flow)
        {
            var list = new List<AudioDeviceInfo>();
            try
            {
                var enumerator = CreateEnumerator();

                // EnumAudioEndpoints is declared with out IntPtr because
                // IMMDeviceCollection doesn't respond to QI from managed code.
                IntPtr colPtr = IntPtr.Zero;
                int hr = enumerator.EnumAudioEndpoints(flow, DeviceState.Active, out colPtr);
                if (hr != 0 || colPtr == IntPtr.Zero) return list.ToArray();

                uint count = CollGetCount(colPtr);

                for (uint i = 0; i < count; i++)
                {
                    IntPtr devPtr = CollItem(colPtr, i);
                    if (devPtr == IntPtr.Zero) continue;
                    try
                    {
                        // Read the device ID via vtable (no QI needed)
                        string id = DevGetIdRaw(devPtr);
                        if (string.IsNullOrEmpty(id)) continue;

                        // Re-fetch through GetDevice so we have a proper
                        // typed IMMDevice for Activate / OpenPropertyStore
                        IMMDevice dev = null;
                        enumerator.GetDevice(id, out dev);
                        if (dev == null) continue;

                        // Friendly name
                        IPropertyStore props = null;
                        dev.OpenPropertyStore(0 /*STGM_READ*/, out props);
                        var        key = PKEY_FriendlyName;
                        PropVariant pv;
                        props.GetValue(ref key, out pv);
                        string name = pv.GetStringValue();
                        if (string.IsNullOrEmpty(name)) name = id;

                        // Volume + mute
                        var    iid    = IID_IAudioEndpointVolume;
                        object o      = null;
                        dev.Activate(ref iid, 23, IntPtr.Zero, out o);
                        var   vol     = (IAudioEndpointVolume)o;
                        float scalar  = 0;
                        bool  muted   = false;
                        vol.GetMasterVolumeLevelScalar(out scalar);
                        vol.GetMute(out muted);

                        list.Add(new AudioDeviceInfo
                        {
                            DeviceId     = id,
                            Name         = name,
                            VolumeScalar = scalar,
                            IsMuted      = muted
                        });
                    }
                    catch { /* skip bad device */ }
                }
            }
            catch { }
            return list.ToArray();
        }

        // ---- device volume / mute -----------------------------------------

        public static bool SetDeviceVolume(string deviceId, float level)
        {
            return RunOnSTA(() => {
                try
                {
                    var  vol  = GetEndpointVolume(deviceId);
                    var  guid = Guid.Empty;
                    vol.SetMasterVolumeLevelScalar(level, ref guid);
                    return true;
                }
                catch { return false; }
            });
        }

        public static bool SetDeviceMute(string deviceId, bool muted)
        {
            return RunOnSTA(() => {
                try
                {
                    var vol  = GetEndpointVolume(deviceId);
                    var guid = Guid.Empty;
                    vol.SetMute(muted, ref guid);
                    return true;
                }
                catch { return false; }
            });
        }

        // ---- default device -----------------------------------------------

        public static bool SetDefaultDevice(string deviceId)
        {
            return RunOnSTA(() => {
                try
                {
                    var policy = CreatePolicyConfig();
                    foreach (ERole role in new[] { ERole.eConsole, ERole.eMultimedia, ERole.eCommunications })
                        policy.SetDefaultEndpoint(deviceId, role);
                    return true;
                }
                catch { return false; }
            });
        }

        // ---- audio sessions -----------------------------------------------

        public static AudioSessionInfo[] GetAudioSessions()
        {
            return RunOnSTA(() => GetAudioSessionsCore());
        }

        private static AudioSessionInfo[] GetAudioSessionsCore()
        {
            var list = new List<AudioSessionInfo>();
            try
            {
                var enumerator = CreateEnumerator();
                IMMDevice defaultDev = null;
                enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eConsole, out defaultDev);

                var  iid2 = IID_IAudioSessionManager2;
                object o2 = null;
                defaultDev.Activate(ref iid2, 23, IntPtr.Zero, out o2);
                var mgr2 = (IAudioSessionManager2)o2;

                IAudioSessionEnumerator sessionEnum = null;
                mgr2.GetSessionEnumerator(out sessionEnum);
                int count = 0;
                sessionEnum.GetCount(out count);

                for (int i = 0; i < count; i++)
                {
                    IAudioSessionControl ctrl = null;
                    sessionEnum.GetSession(i, out ctrl);
                    try
                    {
                        var ctrl2     = (IAudioSessionControl2)ctrl;
                        var simpleVol = (ISimpleAudioVolume)ctrl;

                        uint pid = 0;
                        ctrl2.GetProcessId(out pid);
                        if (pid == 0) continue;

                        string displayName = null;
                        ctrl2.GetDisplayName(out displayName);

                        if (string.IsNullOrEmpty(displayName))
                        {
                            try
                            {
                                var proc = Process.GetProcessById((int)pid);
                                displayName = string.IsNullOrEmpty(proc.MainWindowTitle)
                                    ? proc.ProcessName
                                    : proc.MainWindowTitle;
                            }
                            catch { displayName = "PID " + pid; }
                        }

                        float volLevel = 0;
                        bool  muted    = false;
                        simpleVol.GetMasterVolume(out volLevel);
                        simpleVol.GetMute(out muted);

                        AudioSessionState state = AudioSessionState.Inactive;
                        ctrl.GetState(out state);

                        int pct = (int)Math.Round(volLevel * 100);

                        list.Add(new AudioSessionInfo
                        {
                            SessionKey    = pid + "-" + i,
                            Name          = displayName,
                            ProcessId     = (int)pid,
                            PidLabel      = "PID: " + pid,
                            VolumePercent = pct,
                            VolumeLabel   = pct + "%",
                            IsMuted       = muted,
                            State         = state.ToString(),
                            Icon          = GetProcessIcon((int)pid)
                        });
                    }
                    catch { /* expired session */ }
                }
            }
            catch { }
            return list.ToArray();
        }

        private static string GetProcessIcon(int pid)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "chrome",             "[web]"    },
                { "firefox",            "[web]"    },
                { "msedge",             "[web]"    },
                { "spotify",            "[music]"  },
                { "discord",            "[chat]"   },
                { "slack",              "[chat]"   },
                { "teams",              "[meet]"   },
                { "zoom",               "[meet]"   },
                { "vlc",                "[video]"  },
                { "obs64",              "[stream]" },
                { "obs32",              "[stream]" },
                { "steam",              "[game]"   },
                { "epicgameslauncher",  "[game]"   },
                { "mpc-hc64",           "[video]"  },
                { "foobar2000",         "[music]"  },
                { "itunes",             "[music]"  },
                { "winamp",             "[music]"  }
            };
            try
            {
                var proc = Process.GetProcessById(pid);
                string val;
                if (map.TryGetValue(proc.ProcessName, out val)) return val;
            }
            catch { }
            return "[audio]";
        }

        // ---- per-app session volume / mute --------------------------------
        // Looks up the session by ProcessId each call - reliable and simple.

        public static bool SetSessionVolume(int processId, float level)
        {
            return RunOnSTA(() => SetSessionParam(processId, level, null));
        }

        public static bool SetSessionMute(int processId, bool muted)
        {
            return RunOnSTA(() => SetSessionParam(processId, null, muted));
        }

        private static bool SetSessionParam(int processId, float? volume, bool? muted)
        {
            try
            {
                var enumerator = CreateEnumerator();
                IMMDevice defaultDev = null;
                enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eConsole, out defaultDev);

                var  iid = IID_IAudioSessionManager2;
                object o = null;
                defaultDev.Activate(ref iid, 23, IntPtr.Zero, out o);
                var mgr2 = (IAudioSessionManager2)o;

                IAudioSessionEnumerator sessionEnum = null;
                mgr2.GetSessionEnumerator(out sessionEnum);
                int count = 0;
                sessionEnum.GetCount(out count);

                for (int i = 0; i < count; i++)
                {
                    IAudioSessionControl ctrl = null;
                    sessionEnum.GetSession(i, out ctrl);
                    try
                    {
                        var ctrl2 = (IAudioSessionControl2)ctrl;
                        uint pid  = 0;
                        ctrl2.GetProcessId(out pid);
                        if ((int)pid != processId) continue;

                        var  sv   = (ISimpleAudioVolume)ctrl;
                        var  guid = Guid.Empty;
                        if (volume.HasValue)  sv.SetMasterVolume(volume.Value, ref guid);
                        if (muted.HasValue)   sv.SetMute(muted.Value, ref guid);
                        return true;
                    }
                    catch { }
                }
                return false;
            }
            catch { return false; }
        }

        // ---- device format ------------------------------------------------

        public static DeviceFormatInfo GetDeviceFormat(string deviceId)
        {
            return RunOnSTA<DeviceFormatInfo>(() => {
                try
                {
                    var policy = CreatePolicyConfig();
                    IntPtr fmtPtr = IntPtr.Zero;
                    policy.GetDeviceFormat(deviceId, false, out fmtPtr);
                    if (fmtPtr == IntPtr.Zero) return null;

                    var wfx = Marshal.PtrToStructure<WAVEFORMATEX>(fmtPtr);
                    Marshal.FreeCoTaskMem(fmtPtr);

                    return new DeviceFormatInfo
                    {
                        DeviceId   = deviceId,
                        SampleRate = (int)wfx.nSamplesPerSec,
                        BitDepth   = wfx.wBitsPerSample,
                        Channels   = wfx.nChannels,
                        Label      = wfx.nSamplesPerSec + " Hz / " + wfx.wBitsPerSample + "-bit / "
                                     + (wfx.nChannels == 1 ? "Mono" : wfx.nChannels == 2 ? "Stereo"
                                        : wfx.nChannels + "ch")
                    };
                }
                catch { return null; }
            });
        }

        public static bool SetDeviceFormat(string deviceId, int sampleRate, int bitDepth, int channels)
        {
            return RunOnSTA(() => {
                try
                {
                    var wfx = new WAVEFORMATEX
                    {
                        wFormatTag      = 1, // WAVE_FORMAT_PCM
                        nChannels       = (ushort)channels,
                        nSamplesPerSec  = (uint)sampleRate,
                        wBitsPerSample  = (ushort)bitDepth,
                        nBlockAlign     = (ushort)(channels * bitDepth / 8),
                        nAvgBytesPerSec = (uint)(sampleRate * channels * bitDepth / 8),
                        cbSize          = 0
                    };

                    // For 24-bit or 32-bit, use WAVE_FORMAT_EXTENSIBLE (0xFFFE)
                    // For 16-bit PCM keep WAVE_FORMAT_PCM (1)
                    IntPtr fmtPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(wfx));
                    Marshal.StructureToPtr(wfx, fmtPtr, false);
                    var policy = CreatePolicyConfig();
                    policy.SetDeviceFormat(deviceId, fmtPtr, IntPtr.Zero);
                    Marshal.FreeCoTaskMem(fmtPtr);
                    return true;
                }
                catch { return false; }
            });
        }
    }
}

'@

try {
    Add-Type -TypeDefinition $script:CoreAudioCSharp -Language CSharp -ReferencedAssemblies @(
        'System',
        'System.Runtime.InteropServices'
    ) -ErrorAction Stop
} catch {
    Write-Warning "CoreAudio type already loaded or failed to load: $_"
}


function Get-AudioDevices {
    $result = @{ Render = @(); Capture = @() }
    try {
        $result.Render  = [AudioManager.AudioManagerHelper]::GetRenderDevices()
        $result.Capture = [AudioManager.AudioManagerHelper]::GetCaptureDevices()
    } catch {
        Write-Warning "Get-AudioDevices error: $_"
    }
    return $result
}


function Get-AudioEnhancement {
    param([Parameter(Mandatory)][string]$DeviceId)
    # DeviceId format: {0.0.0.00000000}.{GUID}  - extract the GUID portion
    $guidPart = ($DeviceId -split '\.')[-1].Trim('{', '}')
    $regPath  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{$guidPart}\FxProperties"

    if (-not (Test-Path $regPath)) {
        # Try capture path
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{$guidPart}\FxProperties"
    }

    try {
        $val = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        # PKEY_AudioEndpoint_Disable_SysFx = {1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5
        $propName = "{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5"
        if ($null -ne $val -and $null -ne $val.$propName) {
            # Value 0 = enhancements enabled, 1 = disabled
            return ($val.$propName -eq 0)
        }
        return $true  # Default: enhancements enabled
    } catch {
        return $true
    }
}


function Get-AudioProfiles {
    try {
        if (Test-Path $sync.ProfilesPath) {
            $json = Get-Content -Path $sync.ProfilesPath -Raw -ErrorAction Stop
            $profiles = $json | ConvertFrom-Json
            return @($profiles)
        }
    } catch {
        Write-Warning "Get-AudioProfiles error: $_"
    }
    return @()
}


function Get-AudioSessions {
    try {
        return [AudioManager.AudioManagerHelper]::GetAudioSessions()
    } catch {
        Write-Warning "Get-AudioSessions error: $_"
        return @()
    }
}


function Get-DeviceFormat {
    param([Parameter(Mandatory)][string]$DeviceId)
    try {
        return [AudioManager.AudioManagerHelper]::GetDeviceFormat($DeviceId)
    } catch {
        Write-Warning "Get-DeviceFormat error: $_"
        return $null
    }
}


function Get-DeviceMute {
    param([Parameter(Mandatory)][string]$DeviceId)
    $devices = [AudioManager.AudioManagerHelper]::GetRenderDevices() +
               [AudioManager.AudioManagerHelper]::GetCaptureDevices()
    $dev = $devices | Where-Object { $_.DeviceId -eq $DeviceId } | Select-Object -First 1
    return if ($dev) { $dev.IsMuted } else { $false }
}


function Get-DeviceVolume {
    param([Parameter(Mandatory)][string]$DeviceId)
    $devices = [AudioManager.AudioManagerHelper]::GetRenderDevices() +
               [AudioManager.AudioManagerHelper]::GetCaptureDevices()
    $dev = $devices | Where-Object { $_.DeviceId -eq $DeviceId } | Select-Object -First 1
    return if ($dev) { $dev.VolumeScalar } else { 0.0 }
}


function Invoke-AudioManagerRunspace {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock)
    $ps = [powershell]::Create()
    $ps.RunspacePool = $sync.RunspacePool
    $ps.AddScript($ScriptBlock) | Out-Null
    $ps.BeginInvoke() | Out-Null
}


function Remove-AudioProfile {
    param([Parameter(Mandatory)]$Profile)
    try {
        $existing = Get-AudioProfiles
        $existing  = @($existing | Where-Object { $_.name -ne $Profile.name })
        $existing | ConvertTo-Json -Depth 5 | Set-Content -Path $sync.ProfilesPath -Encoding UTF8
        $sync.Profiles = Get-AudioProfiles
        Set-WPFStatus "Profile '$($Profile.name)' deleted."
    } catch {
        Write-Warning "Remove-AudioProfile error: $_"
    }
}


function Restore-AudioProfile {
    param([Parameter(Mandatory)]$Profile)
    try {
        # Restore default output device
        if ($Profile.defaultOutputDeviceId) {
            Set-DefaultAudioDevice -DeviceId $Profile.defaultOutputDeviceId -Flow ([AudioManager.EDataFlow]::eRender)
        }

        # Restore default input device
        if ($Profile.defaultInputDeviceId) {
            Set-DefaultAudioDevice -DeviceId $Profile.defaultInputDeviceId -Flow ([AudioManager.EDataFlow]::eCapture)
        }

        # Restore output volume
        if ($null -ne $Profile.outputVolume -and $Profile.defaultOutputDeviceId) {
            Set-DeviceVolume -DeviceId $Profile.defaultOutputDeviceId -Level $Profile.outputVolume
            Set-DeviceMute   -DeviceId $Profile.defaultOutputDeviceId -Muted  $Profile.outputMuted
        }

        # Restore input volume
        if ($null -ne $Profile.inputVolume -and $Profile.defaultInputDeviceId) {
            Set-DeviceVolume -DeviceId $Profile.defaultInputDeviceId -Level $Profile.inputVolume
            Set-DeviceMute   -DeviceId $Profile.defaultInputDeviceId -Muted  $Profile.inputMuted
        }

        # Restore per-app volumes by matching process names
        if ($Profile.appVolumes -and $sync.AudioSessions.Count -gt 0) {
            foreach ($saved in $Profile.appVolumes) {
                $live = $sync.AudioSessions | Where-Object { $_.Name -like "*$($saved.processName)*" } | Select-Object -First 1
                if ($live) {
                    Set-AppVolume -SessionKey $live.SessionKey -Level $saved.volume
                    Set-AppMute   -SessionKey $live.SessionKey -Muted  $saved.muted
                }
            }
        }

        Set-WPFStatus "Profile '$($Profile.name)' restored."
        return $true
    } catch {
        Write-Warning "Restore-AudioProfile error: $_"
        Set-WPFStatus "Failed to restore profile: $_"
        return $false
    }
}


function Save-AudioProfile {
    param([Parameter(Mandatory)][string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        Set-WPFStatus "Profile name cannot be empty."
        return
    }

    $profile = [ordered]@{
        name    = $Name.Trim()
        created = (Get-Date -Format 'o')
    }

    $enumerator = [AudioManager.AudioManagerHelper]::CreateDeviceEnumerator()

    # Output device
    if ($sync.WPFProfileSaveOutputDevice.IsChecked) {
        try {
            $dev = $null
            $enumerator.GetDefaultAudioEndpoint(
                [AudioManager.EDataFlow]::eRender,
                [AudioManager.ERole]::eConsole,
                [ref]$dev
            ) | Out-Null
            $id = ""
            $dev.GetId([ref]$id) | Out-Null
            $profile.defaultOutputDeviceId   = $id
            $profile.defaultOutputDeviceName = ($sync.RenderDevices | Where-Object { $_.DeviceId -eq $id } | Select-Object -First 1).Name
        } catch {}
    }

    # Input device
    if ($sync.WPFProfileSaveInputDevice.IsChecked) {
        try {
            $dev = $null
            $enumerator.GetDefaultAudioEndpoint(
                [AudioManager.EDataFlow]::eCapture,
                [AudioManager.ERole]::eConsole,
                [ref]$dev
            ) | Out-Null
            $id = ""
            $dev.GetId([ref]$id) | Out-Null
            $profile.defaultInputDeviceId   = $id
            $profile.defaultInputDeviceName = ($sync.CaptureDevices | Where-Object { $_.DeviceId -eq $id } | Select-Object -First 1).Name
        } catch {}
    }

    # Output volume
    if ($sync.WPFProfileSaveOutputVolume.IsChecked -and $sync.SelectedOutputId) {
        $profile.outputVolume = Get-DeviceVolume -DeviceId $sync.SelectedOutputId
        $profile.outputMuted  = Get-DeviceMute   -DeviceId $sync.SelectedOutputId
    }

    # Input volume
    if ($sync.WPFProfileSaveInputVolume.IsChecked -and $sync.SelectedInputId) {
        $profile.inputVolume = Get-DeviceVolume -DeviceId $sync.SelectedInputId
        $profile.inputMuted  = Get-DeviceMute   -DeviceId $sync.SelectedInputId
    }

    # Per-app volumes
    if ($sync.WPFProfileSaveAppVolumes.IsChecked -and $sync.AudioSessions.Count -gt 0) {
        $appVols = @()
        foreach ($s in $sync.AudioSessions) {
            $appVols += @{ processName = $s.Name; volume = ($s.VolumePercent / 100.0); muted = $s.IsMuted }
        }
        $profile.appVolumes = $appVols
    }

    # Load existing profiles, remove duplicate name, append new
    $existing = Get-AudioProfiles
    $existing  = @($existing | Where-Object { $_.name -ne $profile.name })
    $existing += [PSCustomObject]$profile

    $existing | ConvertTo-Json -Depth 5 | Set-Content -Path $sync.ProfilesPath -Encoding UTF8
    $sync.Profiles = Get-AudioProfiles
    Set-WPFStatus "Profile '$Name' saved."
}


function Set-AppMute {
    param(
        [Parameter(Mandatory)][int]$ProcessId,
        [Parameter(Mandatory)][bool]$Muted
    )
    return [AudioManager.AudioManagerHelper]::SetSessionMute($ProcessId, $Muted)
}


function Set-AppVolume {
    param(
        [Parameter(Mandatory)][int]$ProcessId,
        [Parameter(Mandatory)][float]$Level    # 0.0 - 1.0
    )
    $Level = [math]::Max(0.0, [math]::Min(1.0, $Level))
    return [AudioManager.AudioManagerHelper]::SetSessionVolume($ProcessId, $Level)
}


function Set-AudioEnhancement {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][bool]$Enabled
    )
    $guidPart = ($DeviceId -split '\.')[-1].Trim('{', '}')
    $regPath  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{$guidPart}\FxProperties"

    if (-not (Test-Path $regPath)) {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{$guidPart}\FxProperties"
    }

    try {
        $propName = "{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5"
        $value    = if ($Enabled) { 0 } else { 1 }  # 0 = enabled, 1 = disabled

        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name $propName -Value $value -Type DWord -Force

        # Restart audio service to apply changes
        Restart-Service -Name audiosrv -Force -ErrorAction SilentlyContinue
        return $true
    } catch {
        Write-Warning "Set-AudioEnhancement error: $_"
        return $false
    }
}


function Set-DefaultAudioDevice {
    param([Parameter(Mandatory)][string]$DeviceId)
    return [AudioManager.AudioManagerHelper]::SetDefaultDevice($DeviceId)
}


function Set-DeviceFormat {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][int]$SampleRate,
        [Parameter(Mandatory)][int]$BitDepth,
        [int]$Channels = 2
    )
    return [AudioManager.AudioManagerHelper]::SetDeviceFormat($DeviceId, $SampleRate, $BitDepth, $Channels)
}


function Set-DeviceMute {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][bool]$Muted
    )
    return [AudioManager.AudioManagerHelper]::SetDeviceMute($DeviceId, $Muted)
}


function Set-DeviceVolume {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][float]$Level     # 0.0 - 1.0
    )
    $Level = [math]::Max(0.0, [math]::Min(1.0, $Level))
    return [AudioManager.AudioManagerHelper]::SetDeviceVolume($DeviceId, $Level)
}


function Initialize-ApplicationsTab {
    $sync.WPFAppSessionList.Children.Clear()

    $sessions = $sync.AudioSessions
    if (-not $sessions -or $sessions.Count -eq 0) {
        $sync.WPFAppCount.Text = "No active audio sessions found."
        return
    }

    $sync.WPFAppCount.Text = "$($sessions.Count) session(s)"

    $bgColor  = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString("#16213E")
    $fgMain   = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString("#E0E0E0")
    $fgMuted  = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString("#8892B0")

    foreach ($session in ($sessions | Sort-Object Name)) {
        # --- outer card border ---
        $border = [System.Windows.Controls.Border]::new()
        $border.Background   = $bgColor
        $border.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $border.Padding      = [System.Windows.Thickness]::new(12, 8, 12, 8)
        $border.Margin       = [System.Windows.Thickness]::new(0, 0, 0, 8)

        # --- 5-column grid ---
        $grid = [System.Windows.Controls.Grid]::new()
        foreach ($w in @(36, 180, -1, 50, 70)) {
            $cd = [System.Windows.Controls.ColumnDefinition]::new()
            $cd.Width = if ($w -lt 0) { [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) } `
                        else           { [System.Windows.GridLength]::new($w) }
            $grid.ColumnDefinitions.Add($cd)
        }

        # Col 0: icon
        $iconTb = [System.Windows.Controls.TextBlock]::new()
        $iconTb.Text                = $session.Icon
        $iconTb.FontSize            = 14
        $iconTb.VerticalAlignment   = "Center"
        $iconTb.HorizontalAlignment = "Center"
        [System.Windows.Controls.Grid]::SetColumn($iconTb, 0)

        # Col 1: name + pid
        $namePanel = [System.Windows.Controls.StackPanel]::new()
        $namePanel.VerticalAlignment = "Center"
        $namePanel.Margin = [System.Windows.Thickness]::new(8, 0, 0, 0)

        $nameTb = [System.Windows.Controls.TextBlock]::new()
        $nameTb.Text            = $session.Name
        $nameTb.FontWeight      = "SemiBold"
        $nameTb.Foreground      = $fgMain
        $nameTb.TextTrimming    = "CharacterEllipsis"

        $pidTb = [System.Windows.Controls.TextBlock]::new()
        $pidTb.Text       = $session.PidLabel
        $pidTb.Foreground = $fgMuted
        $pidTb.FontSize   = 10

        $namePanel.Children.Add($nameTb) | Out-Null
        $namePanel.Children.Add($pidTb)  | Out-Null
        [System.Windows.Controls.Grid]::SetColumn($namePanel, 1)

        # Col 2: volume slider
        $slider = [System.Windows.Controls.Slider]::new()
        $slider.Minimum           = 0
        $slider.Maximum           = 100
        $slider.Value             = $session.VolumePercent
        $slider.VerticalAlignment = "Center"
        $slider.Margin            = [System.Windows.Thickness]::new(8, 0, 8, 0)
        $slider.Tag               = [int]$session.ProcessId
        [System.Windows.Controls.Grid]::SetColumn($slider, 2)

        # Col 3: volume label (updated live by slider)
        $volLabel = [System.Windows.Controls.TextBlock]::new()
        $volLabel.Text                = $session.VolumeLabel
        $volLabel.Foreground          = $fgMain
        $volLabel.FontWeight          = "SemiBold"
        $volLabel.VerticalAlignment   = "Center"
        $volLabel.HorizontalAlignment = "Center"
        [System.Windows.Controls.Grid]::SetColumn($volLabel, 3)

        # Col 4: mute toggle
        $muteBtn = [System.Windows.Controls.Primitives.ToggleButton]::new()
        $muteBtn.Content          = if ($session.IsMuted) { "[X]" } else { "Mute" }
        $muteBtn.IsChecked        = $session.IsMuted
        $muteBtn.VerticalAlignment = "Center"
        $muteBtn.Tag              = [int]$session.ProcessId
        $muteBtn.Padding          = [System.Windows.Thickness]::new(8, 4, 8, 4)
        [System.Windows.Controls.Grid]::SetColumn($muteBtn, 4)

        # --- events ---
        $capturedLabel = $volLabel
        $slider.Add_ValueChanged({
            $capturedLabel.Text = "$([math]::Round($this.Value))%"
        }.GetNewClosure())

        $slider.Add_PreviewMouseLeftButtonUp({
            $pid   = [int]$this.Tag
            $level = [float]($this.Value / 100.0)
            $sync._tmpPid   = $pid
            $sync._tmpLevel = $level
            Invoke-AudioManagerRunspace { Set-AppVolume -ProcessId $sync._tmpPid -Level $sync._tmpLevel }
        }.GetNewClosure())

        $muteBtn.Add_Click({
            $pid   = [int]$this.Tag
            $muted = [bool]$this.IsChecked
            $this.Content = if ($muted) { "[X]" } else { "Mute" }
            $sync._tmpPid   = $pid
            $sync._tmpMuted = $muted
            Invoke-AudioManagerRunspace { Set-AppMute -ProcessId $sync._tmpPid -Muted $sync._tmpMuted }
        }.GetNewClosure())

        # --- assemble ---
        $grid.Children.Add($iconTb)    | Out-Null
        $grid.Children.Add($namePanel) | Out-Null
        $grid.Children.Add($slider)    | Out-Null
        $grid.Children.Add($volLabel)  | Out-Null
        $grid.Children.Add($muteBtn)   | Out-Null
        $border.Child = $grid
        $sync.WPFAppSessionList.Children.Add($border) | Out-Null
    }
}

function Update-ApplicationsTab {
    Initialize-ApplicationsTab
}


function Initialize-DevicesTab {
    # Populate output device list
    $sync.WPFOutputDeviceList.Items.Clear()
    foreach ($dev in $sync.RenderDevices) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = $dev.Name
        $item.Tag     = $dev.DeviceId
        $sync.WPFOutputDeviceList.Items.Add($item) | Out-Null
    }

    # Populate input device list
    $sync.WPFInputDeviceList.Items.Clear()
    foreach ($dev in $sync.CaptureDevices) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = $dev.Name
        $item.Tag     = $dev.DeviceId
        $sync.WPFInputDeviceList.Items.Add($item) | Out-Null
    }
}

function Update-DevicesTab {
    Initialize-DevicesTab

    # If a device was selected, restore selection and update sliders
    if ($sync.SelectedOutputId) {
        $items = $sync.WPFOutputDeviceList.Items
        for ($i = 0; $i -lt $items.Count; $i++) {
            if ($items[$i].Tag -eq $sync.SelectedOutputId) {
                $sync.WPFOutputDeviceList.SelectedIndex = $i
                break
            }
        }
    }

    if ($sync.SelectedInputId) {
        $items = $sync.WPFInputDeviceList.Items
        for ($i = 0; $i -lt $items.Count; $i++) {
            if ($items[$i].Tag -eq $sync.SelectedInputId) {
                $sync.WPFInputDeviceList.SelectedIndex = $i
                break
            }
        }
    }
}


function Initialize-FormatsTab {
    # Populate output device picker
    $sync.WPFFormatOutputDevice.Items.Clear()
    foreach ($dev in $sync.RenderDevices) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $dev.Name
        $item.Tag     = $dev.DeviceId
        $sync.WPFFormatOutputDevice.Items.Add($item) | Out-Null
    }
    if ($sync.WPFFormatOutputDevice.Items.Count -gt 0) {
        $sync.WPFFormatOutputDevice.SelectedIndex = 0
        Update-OutputFormatDisplay
    }

    # Populate input device picker
    $sync.WPFFormatInputDevice.Items.Clear()
    foreach ($dev in $sync.CaptureDevices) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $dev.Name
        $item.Tag     = $dev.DeviceId
        $sync.WPFFormatInputDevice.Items.Add($item) | Out-Null
    }
    if ($sync.WPFFormatInputDevice.Items.Count -gt 0) {
        $sync.WPFFormatInputDevice.SelectedIndex = 0
        Update-InputFormatDisplay
    }
}

function Update-OutputFormatDisplay {
    $selected = $sync.WPFFormatOutputDevice.SelectedItem
    if (-not $selected) { return }
    $deviceId = $selected.Tag
    $fmt = Get-DeviceFormat -DeviceId $deviceId
    if ($fmt) {
        $sync.WPFCurrentOutputFormat.Text = $fmt.Description
        $sync.WPFApplyOutputFormat.IsEnabled = $true

        # Reflect enhancements state
        $enhanced = Get-AudioEnhancement -DeviceId $deviceId
        $sync.WPFOutputEnhancementsToggle.IsChecked = $enhanced
        $sync.WPFOutputEnhancementsToggle.Content   = if ($enhanced) { "Enhancements: Enabled" } else { "Enhancements: Disabled" }
        $sync.WPFOutputEnhancementsToggle.IsEnabled  = $true
    } else {
        $sync.WPFCurrentOutputFormat.Text = "Unable to read format"
    }
}

function Update-InputFormatDisplay {
    $selected = $sync.WPFFormatInputDevice.SelectedItem
    if (-not $selected) { return }
    $deviceId = $selected.Tag
    $fmt = Get-DeviceFormat -DeviceId $deviceId
    if ($fmt) {
        $sync.WPFCurrentInputFormat.Text = $fmt.Description
        $sync.WPFApplyInputFormat.IsEnabled = $true

        $enhanced = Get-AudioEnhancement -DeviceId $deviceId
        $sync.WPFInputEnhancementsToggle.IsChecked = $enhanced
        $sync.WPFInputEnhancementsToggle.Content   = if ($enhanced) { "Enhancements: Enabled" } else { "Enhancements: Disabled" }
        $sync.WPFInputEnhancementsToggle.IsEnabled  = $true
    } else {
        $sync.WPFCurrentInputFormat.Text = "Unable to read format"
    }
}


function Initialize-ProfilesTab {
    $sync.WPFProfileList.Items.Clear()
    $sync.WPFRestoreProfile.IsEnabled = $false
    $sync.WPFDeleteProfile.IsEnabled  = $false
    $sync.WPFProfileInfo.Text         = ""

    if ($sync.Profiles.Count -eq 0) {
        $sync.WPFProfileList.Items.Add("No saved profiles") | Out-Null
        return
    }

    foreach ($p in $sync.Profiles) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = $p.name
        $item.Tag     = $p
        $item.ToolTip = "Saved: $($p.created)"
        $sync.WPFProfileList.Items.Add($item) | Out-Null
    }
}

function Update-ProfilesTab {
    Initialize-ProfilesTab
}


function Invoke-WPFButton {
    param([Parameter(Mandatory)][string]$ClickedButton)

    switch ($ClickedButton) {

        "WPFRefreshButton" {
            Invoke-WPFRefreshAll
        }

        "WPFRefreshApps" {
            Set-WPFStatus "Refreshing application sessions..."
            Invoke-AudioManagerRunspace {
                $sync.AudioSessions = Get-AudioSessions
                Invoke-WPFUIThread { Update-ApplicationsTab }
                Invoke-WPFUIThread { Set-WPFStatus "Found $($sync.AudioSessions.Count) active session(s)." }
            }
        }

        "WPFSetDefaultOutput" {
            $selected = $sync.WPFOutputDeviceList.SelectedItem
            if (-not $selected) { Set-WPFStatus "Select an output device first."; return }
            $deviceId = $selected.Tag
            $devName  = $selected.Content
            Set-WPFStatus "Setting default output to '$devName'..."
            Invoke-AudioManagerRunspace {
                $ok = Set-DefaultAudioDevice -DeviceId $deviceId
                Invoke-WPFUIThread {
                    Set-WPFStatus (if ($ok) { "Default output set to '$devName'." } else { "Failed to set default output." })
                }
            }
        }

        "WPFSetDefaultInput" {
            $selected = $sync.WPFInputDeviceList.SelectedItem
            if (-not $selected) { Set-WPFStatus "Select an input device first."; return }
            $deviceId = $selected.Tag
            $devName  = $selected.Content
            Set-WPFStatus "Setting default input to '$devName'..."
            Invoke-AudioManagerRunspace {
                $ok = Set-DefaultAudioDevice -DeviceId $deviceId
                Invoke-WPFUIThread {
                    Set-WPFStatus (if ($ok) { "Default input set to '$devName'." } else { "Failed to set default input." })
                }
            }
        }

        "WPFApplyOutputFormat" {
            $selected = $sync.WPFFormatOutputDevice.SelectedItem
            if (-not $selected) { Set-WPFStatus "Select an output device first."; return }
            $deviceId  = $selected.Tag
            $srItem    = $sync.WPFOutputSampleRate.SelectedItem
            $bdItem    = $sync.WPFOutputBitDepth.SelectedItem
            if (-not $srItem -or -not $bdItem) { Set-WPFStatus "Select sample rate and bit depth."; return }
            $sampleRate = [uint32]$srItem.Tag
            $bitDepth   = [uint16]$bdItem.Tag
            Set-WPFStatus "Applying output format: ${sampleRate} Hz / ${bitDepth}-bit..."
            Invoke-AudioManagerRunspace {
                $ok = Set-DeviceFormat -DeviceId $deviceId -SampleRate $sampleRate -BitDepth $bitDepth
                Invoke-WPFUIThread {
                    if ($ok) {
                        Update-OutputFormatDisplay
                        Set-WPFStatus "Output format applied."
                    } else {
                        Set-WPFStatus "Failed to apply output format."
                    }
                }
            }
        }

        "WPFApplyInputFormat" {
            $selected = $sync.WPFFormatInputDevice.SelectedItem
            if (-not $selected) { Set-WPFStatus "Select an input device first."; return }
            $deviceId  = $selected.Tag
            $srItem    = $sync.WPFInputSampleRate.SelectedItem
            $bdItem    = $sync.WPFInputBitDepth.SelectedItem
            if (-not $srItem -or -not $bdItem) { Set-WPFStatus "Select sample rate and bit depth."; return }
            $sampleRate = [uint32]$srItem.Tag
            $bitDepth   = [uint16]$bdItem.Tag
            Set-WPFStatus "Applying input format: ${sampleRate} Hz / ${bitDepth}-bit..."
            Invoke-AudioManagerRunspace {
                $ok = Set-DeviceFormat -DeviceId $deviceId -SampleRate $sampleRate -BitDepth $bitDepth
                Invoke-WPFUIThread {
                    if ($ok) {
                        Update-InputFormatDisplay
                        Set-WPFStatus "Input format applied."
                    } else {
                        Set-WPFStatus "Failed to apply input format."
                    }
                }
            }
        }

        "WPFSaveProfile" {
            $name = $sync.WPFProfileNameInput.Text
            Save-AudioProfile -Name $name
            Invoke-WPFUIThread { Update-ProfilesTab }
        }

        "WPFRestoreProfile" {
            $selected = $sync.WPFProfileList.SelectedItem
            if (-not $selected -or -not $selected.Tag) { Set-WPFStatus "Select a profile to restore."; return }
            $profile = $selected.Tag
            Set-WPFStatus "Restoring profile '$($profile.name)'..."
            Invoke-AudioManagerRunspace {
                $ok = Restore-AudioProfile -Profile $profile
                if ($ok) {
                    # Re-read devices after potential device switch
                    $devices = Get-AudioDevices
                    $sync.RenderDevices  = $devices.Render
                    $sync.CaptureDevices = $devices.Capture
                    Invoke-WPFUIThread { Invoke-WPFRefreshAll }
                }
            }
        }

        "WPFDeleteProfile" {
            $selected = $sync.WPFProfileList.SelectedItem
            if (-not $selected -or -not $selected.Tag) { Set-WPFStatus "Select a profile to delete."; return }
            Remove-AudioProfile -Profile $selected.Tag
            Invoke-WPFUIThread { Update-ProfilesTab }
        }
    }
}


function Invoke-WPFRefreshAll {
    if ($sync.IsRefreshing) { return }
    $sync.IsRefreshing = $true
    Set-WPFStatus "Refreshing audio devices..."

    Invoke-AudioManagerRunspace {
        try {
            $devices  = Get-AudioDevices
            $sessions = Get-AudioSessions
            $profiles = Get-AudioProfiles

            $sync.RenderDevices  = $devices.Render
            $sync.CaptureDevices = $devices.Capture
            $sync.AudioSessions  = $sessions
            $sync.Profiles       = $profiles

            Invoke-WPFUIThread {
                Initialize-DevicesTab
                Initialize-ApplicationsTab
                Initialize-FormatsTab
                Initialize-ProfilesTab
                Set-WPFStatus "Ready - $($sync.RenderDevices.Count) output, $($sync.CaptureDevices.Count) input device(s) found."
            }
        } catch {
            Invoke-WPFUIThread { Set-WPFStatus "Refresh error: $_" }
        } finally {
            $sync.IsRefreshing = $false
        }
    }
}


function Invoke-WPFTab {
    param([Parameter(Mandatory)][string]$ClickedTab)

    $tabMap = @{
        WPFTab1BT = 'WPFTab1'
        WPFTab2BT = 'WPFTab2'
        WPFTab3BT = 'WPFTab3'
        WPFTab4BT = 'WPFTab4'
    }

    foreach ($btn in $tabMap.Keys) {
        $sync[$btn].IsChecked = ($btn -eq $ClickedTab)
    }

    $targetTab = $tabMap[$ClickedTab]
    foreach ($tab in $tabMap.Values) {
        $tabItem = $sync.WPFTabControl.Items | Where-Object { $_.Name -eq $tab }
        if ($tabItem) { $tabItem.IsSelected = ($tab -eq $targetTab) }
    }

    $sync.CurrentTab = $targetTab
}


function Invoke-WPFUIThread {
    param([Parameter(Mandatory)][scriptblock]$Code)
    if ($sync.Form -and $sync.Form.Dispatcher) {
        $sync.Form.Dispatcher.Invoke(
            [action]$Code,
            [System.Windows.Threading.DispatcherPriority]::Normal
        )
    }
}

function Set-WPFStatus {
    param([string]$Message)
    Invoke-WPFUIThread {
        if ($sync.WPFStatusBar) {
            $sync.WPFStatusBar.Text = $Message
        }
    }
}


$inputXML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Audio Manager"
    MinWidth="960" MinHeight="680"
    Width="1080" Height="720"
    WindowStartupLocation="CenterScreen"
    Background="#1A1A2E"
    Foreground="#E0E0E0"
    FontFamily="Segoe UI"
    FontSize="13">

    <Window.Resources>
        <SolidColorBrush x:Key="BgDark"      Color="#1A1A2E"/>
        <SolidColorBrush x:Key="BgMid"       Color="#16213E"/>
        <SolidColorBrush x:Key="BgLight"     Color="#0F3460"/>
        <SolidColorBrush x:Key="Accent"      Color="#E94560"/>
        <SolidColorBrush x:Key="AccentHover" Color="#FF6B81"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#E0E0E0"/>
        <SolidColorBrush x:Key="TextMuted"   Color="#8892B0"/>
        <SolidColorBrush x:Key="Green"       Color="#43D97B"/>

        <!-- Base button -->
        <Style x:Key="BtnBase" TargetType="Button">
            <Setter Property="Background"      Value="#0F3460"/>
            <Setter Property="Foreground"      Value="#E0E0E0"/>
            <Setter Property="BorderBrush"     Value="#E94560"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="14,7"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="FontSize"        Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#1a4a80"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#E94560"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Accent button -->
        <Style x:Key="BtnAccent" TargetType="Button" BasedOn="{StaticResource BtnBase}">
            <Setter Property="Background"  Value="#E94560"/>
            <Setter Property="BorderBrush" Value="#E94560"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FF6B81"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Tab toggle button -->
        <Style x:Key="TabBtn" TargetType="ToggleButton">
            <Setter Property="Background"      Value="Transparent"/>
            <Setter Property="Foreground"      Value="#8892B0"/>
            <Setter Property="BorderThickness" Value="0,0,0,2"/>
            <Setter Property="BorderBrush"     Value="Transparent"/>
            <Setter Property="Padding"         Value="20,10"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ToggleButton">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter Property="Foreground"  Value="#E94560"/>
                                <Setter Property="BorderBrush" Value="#E94560"/>
                                <Setter Property="FontWeight"  Value="SemiBold"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="#E0E0E0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Slider -->
        <Style x:Key="AudioSlider" TargetType="Slider">
            <Setter Property="Minimum"           Value="0"/>
            <Setter Property="Maximum"           Value="100"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Slider">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto" MinHeight="{TemplateBinding MinHeight}"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Track Name="PART_Track" Grid.Row="1">
                                <Track.DecreaseRepeatButton>
                                    <RepeatButton Command="Slider.DecreaseLarge">
                                        <RepeatButton.Template>
                                            <ControlTemplate TargetType="RepeatButton">
                                                <Border Height="4" Background="#E94560" CornerRadius="2"/>
                                            </ControlTemplate>
                                        </RepeatButton.Template>
                                    </RepeatButton>
                                </Track.DecreaseRepeatButton>
                                <Track.IncreaseRepeatButton>
                                    <RepeatButton Command="Slider.IncreaseLarge">
                                        <RepeatButton.Template>
                                            <ControlTemplate TargetType="RepeatButton">
                                                <Border Height="4" Background="#2D3561" CornerRadius="2"/>
                                            </ControlTemplate>
                                        </RepeatButton.Template>
                                    </RepeatButton>
                                </Track.IncreaseRepeatButton>
                                <Track.Thumb>
                                    <Thumb>
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Ellipse Width="14" Height="14" Fill="#E94560"
                                                         Stroke="#FF6B81" StrokeThickness="1"/>
                                            </ControlTemplate>
                                        </Thumb.Template>
                                    </Thumb>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Mute toggle button -->
        <Style x:Key="MuteBtn" TargetType="ToggleButton">
            <Setter Property="Background"      Value="#0F3460"/>
            <Setter Property="Foreground"      Value="#E0E0E0"/>
            <Setter Property="BorderBrush"     Value="#2D3561"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="10,5"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="MinWidth"        Value="64"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ToggleButton">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter Property="Background"  Value="#E94560"/>
                                <Setter Property="BorderBrush" Value="#E94560"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="BorderBrush" Value="#E94560"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Card -->
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background"      Value="#16213E"/>
            <Setter Property="BorderBrush"     Value="#2D3561"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius"    Value="8"/>
            <Setter Property="Padding"         Value="16"/>
        </Style>

        <Style TargetType="ListBox">
            <Setter Property="Background"      Value="#0F1729"/>
            <Setter Property="Foreground"      Value="#E0E0E0"/>
            <Setter Property="BorderBrush"     Value="#2D3561"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="Cursor"  Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#E94560"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1a4a80"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ComboBox">
            <Setter Property="Background"      Value="#0F3460"/>
            <Setter Property="Foreground"      Value="#E0E0E0"/>
            <Setter Property="BorderBrush"     Value="#2D3561"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="8,5"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background"      Value="#0F1729"/>
            <Setter Property="Foreground"      Value="#E0E0E0"/>
            <Setter Property="BorderBrush"     Value="#2D3561"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="8,6"/>
            <Setter Property="CaretBrush"      Value="#E94560"/>
        </Style>

        <Style x:Key="SectionHeader" TargetType="TextBlock">
            <Setter Property="FontSize"   Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#E94560"/>
            <Setter Property="Margin"     Value="0,0,0,10"/>
        </Style>

        <Style x:Key="Label" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#8892B0"/>
            <Setter Property="FontSize"   Value="11"/>
            <Setter Property="Margin"     Value="0,0,0,4"/>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#E0E0E0"/>
        </Style>
    </Window.Resources>

    <DockPanel>

        <!-- HEADER -->
        <Border DockPanel.Dock="Top" Background="#16213E" BorderBrush="#2D3561"
                BorderThickness="0,0,0,1" Padding="20,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="[Audio]" FontSize="14" FontWeight="Bold" Foreground="#E94560"
                               Margin="0,0,8,0" VerticalAlignment="Center"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="Audio Manager" FontSize="16" FontWeight="Bold" Foreground="#E0E0E0"/>
                        <TextBlock Name="WPFVersionLabel" Text="v0.0.0" FontSize="10" Foreground="#8892B0"/>
                    </StackPanel>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Center"
                            VerticalAlignment="Center" Margin="20,0">
                    <TextBlock Text="MASTER" FontSize="10" FontWeight="SemiBold" Foreground="#8892B0"
                               VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <ToggleButton Name="WPFMasterMuteButton" Style="{StaticResource MuteBtn}"
                                  Content="Mute" Width="48" Margin="0,0,10,0"
                                  ToolTip="Mute/Unmute master output"/>
                    <Slider Name="WPFMasterVolumeSlider" Style="{StaticResource AudioSlider}"
                            Width="200" VerticalAlignment="Center"/>
                    <TextBlock Name="WPFMasterVolumeLabel" Text="100%" Foreground="#E0E0E0"
                               FontWeight="SemiBold" Width="40" Margin="10,0,0,0" VerticalAlignment="Center"/>
                </StackPanel>

                <Button Name="WPFRefreshButton" Grid.Column="2" Style="{StaticResource BtnBase}"
                        Content="Refresh" VerticalAlignment="Center"/>
            </Grid>
        </Border>

        <!-- TAB NAV -->
        <Border DockPanel.Dock="Top" Background="#16213E" BorderBrush="#2D3561"
                BorderThickness="0,0,0,1">
            <StackPanel Orientation="Horizontal">
                <ToggleButton Name="WPFTab1BT" Style="{StaticResource TabBtn}" Content="Devices"      IsChecked="True"/>
                <ToggleButton Name="WPFTab2BT" Style="{StaticResource TabBtn}" Content="Applications"/>
                <ToggleButton Name="WPFTab3BT" Style="{StaticResource TabBtn}" Content="Formats"/>
                <ToggleButton Name="WPFTab4BT" Style="{StaticResource TabBtn}" Content="Profiles"/>
            </StackPanel>
        </Border>

        <!-- STATUS BAR -->
        <Border DockPanel.Dock="Bottom" Background="#16213E" BorderBrush="#2D3561"
                BorderThickness="0,1,0,0" Padding="16,6">
            <TextBlock Name="WPFStatusBar" Text="Ready" Foreground="#8892B0" FontSize="11"/>
        </Border>

        <!-- TAB CONTENT -->
        <TabControl Name="WPFTabControl" Background="Transparent" BorderThickness="0">
            <TabControl.Resources>
                <Style TargetType="TabItem">
                    <Setter Property="Visibility" Value="Collapsed"/>
                </Style>
            </TabControl.Resources>

            <!-- TAB 1: DEVICES -->
            <TabItem Name="WPFTab1" IsSelected="True">
                <Grid Margin="20">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Output Devices -->
                    <Border Grid.Column="0" Style="{StaticResource Card}">
                        <DockPanel>
                            <TextBlock DockPanel.Dock="Top" Text="Output Devices" Style="{StaticResource SectionHeader}"/>
                            <Grid DockPanel.Dock="Bottom" Margin="0,12,0,0">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Button Name="WPFSetDefaultOutput" Grid.Row="0"
                                        Style="{StaticResource BtnAccent}"
                                        Content="Set as Default Output" Margin="0,0,0,12"
                                        IsEnabled="False"/>
                                <Grid Grid.Row="1">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <Slider Name="WPFOutputVolumeSlider" Grid.Column="0"
                                            Style="{StaticResource AudioSlider}" IsEnabled="False"/>
                                    <TextBlock Name="WPFOutputVolumeLabel" Grid.Column="1"
                                               Text="--%" Foreground="#E0E0E0" FontWeight="SemiBold"
                                               Width="40" Margin="10,0" VerticalAlignment="Center"/>
                                    <ToggleButton Name="WPFOutputMuteButton" Grid.Column="2"
                                                  Style="{StaticResource MuteBtn}" Content="Mute"
                                                  IsEnabled="False" ToolTip="Mute output device"/>
                                </Grid>
                            </Grid>
                            <ListBox Name="WPFOutputDeviceList" Margin="0,0,0,12"/>
                        </DockPanel>
                    </Border>

                    <!-- Input Devices -->
                    <Border Grid.Column="2" Style="{StaticResource Card}">
                        <DockPanel>
                            <TextBlock DockPanel.Dock="Top" Text="Input Devices" Style="{StaticResource SectionHeader}"/>
                            <Grid DockPanel.Dock="Bottom" Margin="0,12,0,0">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Button Name="WPFSetDefaultInput" Grid.Row="0"
                                        Style="{StaticResource BtnAccent}"
                                        Content="Set as Default Input" Margin="0,0,0,12"
                                        IsEnabled="False"/>
                                <Grid Grid.Row="1">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <Slider Name="WPFInputVolumeSlider" Grid.Column="0"
                                            Style="{StaticResource AudioSlider}" IsEnabled="False"/>
                                    <TextBlock Name="WPFInputVolumeLabel" Grid.Column="1"
                                               Text="--%" Foreground="#E0E0E0" FontWeight="SemiBold"
                                               Width="40" Margin="10,0" VerticalAlignment="Center"/>
                                    <ToggleButton Name="WPFInputMuteButton" Grid.Column="2"
                                                  Style="{StaticResource MuteBtn}" Content="Mute"
                                                  IsEnabled="False" ToolTip="Mute input device"/>
                                </Grid>
                            </Grid>
                            <ListBox Name="WPFInputDeviceList" Margin="0,0,0,12"/>
                        </DockPanel>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 2: APPLICATIONS -->
            <TabItem Name="WPFTab2">
                <DockPanel Margin="20">
                    <Border DockPanel.Dock="Top" Style="{StaticResource Card}" Margin="0,0,0,16">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Per-Application Volume" Style="{StaticResource SectionHeader}"
                                       Margin="0" VerticalAlignment="Center"/>
                            <Button Name="WPFRefreshApps" Style="{StaticResource BtnBase}"
                                    Content="Refresh Apps" Margin="16,0,0,0" VerticalAlignment="Center"/>
                            <TextBlock Name="WPFAppCount" Text="" Foreground="#8892B0" FontSize="11"
                                       Margin="12,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                        <StackPanel Name="WPFAppSessionList"/>
                    </ScrollViewer>
                </DockPanel>
            </TabItem>

            <!-- TAB 3: FORMATS -->
            <TabItem Name="WPFTab3">
                <Grid Margin="20">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Output Format -->
                    <Border Grid.Column="0" Style="{StaticResource Card}">
                        <StackPanel>
                            <TextBlock Text="Output Format" Style="{StaticResource SectionHeader}"/>
                            <TextBlock Text="Device" Style="{StaticResource Label}"/>
                            <ComboBox Name="WPFFormatOutputDevice" Margin="0,0,0,12"/>
                            <TextBlock Text="Current Format" Style="{StaticResource Label}"/>
                            <Border Background="#0F1729" BorderBrush="#2D3561" BorderThickness="1"
                                    CornerRadius="4" Padding="8,6" Margin="0,0,0,16">
                                <TextBlock Name="WPFCurrentOutputFormat" Text="--"
                                           Foreground="#43D97B" FontFamily="Consolas" FontSize="12"/>
                            </Border>
                            <TextBlock Text="Sample Rate (Hz)" Style="{StaticResource Label}"/>
                            <ComboBox Name="WPFOutputSampleRate" Margin="0,0,0,12">
                                <ComboBoxItem Content="44100"  Tag="44100"/>
                                <ComboBoxItem Content="48000"  Tag="48000" IsSelected="True"/>
                                <ComboBoxItem Content="88200"  Tag="88200"/>
                                <ComboBoxItem Content="96000"  Tag="96000"/>
                                <ComboBoxItem Content="176400" Tag="176400"/>
                                <ComboBoxItem Content="192000" Tag="192000"/>
                            </ComboBox>
                            <TextBlock Text="Bit Depth" Style="{StaticResource Label}"/>
                            <ComboBox Name="WPFOutputBitDepth" Margin="0,0,0,16">
                                <ComboBoxItem Content="16-bit"        Tag="16"/>
                                <ComboBoxItem Content="24-bit"        Tag="24" IsSelected="True"/>
                                <ComboBoxItem Content="32-bit (float)" Tag="32"/>
                            </ComboBox>
                            <Button Name="WPFApplyOutputFormat" Style="{StaticResource BtnAccent}"
                                    Content="Apply Output Format" Margin="0,0,0,20" IsEnabled="False"/>
                            <Separator Background="#2D3561" Margin="0,0,0,16"/>
                            <TextBlock Text="Audio Enhancements" Style="{StaticResource SectionHeader}"/>
                            <TextBlock Text="Toggle system audio enhancements for this device."
                                       Style="{StaticResource Label}" TextWrapping="Wrap" Margin="0,0,0,10"/>
                            <ToggleButton Name="WPFOutputEnhancementsToggle" Style="{StaticResource MuteBtn}"
                                          Content="Enhancements: ON" MinWidth="180" IsEnabled="False"
                                          HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Border>

                    <!-- Input Format -->
                    <Border Grid.Column="2" Style="{StaticResource Card}">
                        <StackPanel>
                            <TextBlock Text="Input Format" Style="{StaticResource SectionHeader}"/>
                            <TextBlock Text="Device" Style="{StaticResource Label}"/>
                            <ComboBox Name="WPFFormatInputDevice" Margin="0,0,0,12"/>
                            <TextBlock Text="Current Format" Style="{StaticResource Label}"/>
                            <Border Background="#0F1729" BorderBrush="#2D3561" BorderThickness="1"
                                    CornerRadius="4" Padding="8,6" Margin="0,0,0,16">
                                <TextBlock Name="WPFCurrentInputFormat" Text="--"
                                           Foreground="#43D97B" FontFamily="Consolas" FontSize="12"/>
                            </Border>
                            <TextBlock Text="Sample Rate (Hz)" Style="{StaticResource Label}"/>
                            <ComboBox Name="WPFInputSampleRate" Margin="0,0,0,12">
                                <ComboBoxItem Content="8000"  Tag="8000"/>
                                <ComboBoxItem Content="16000" Tag="16000"/>
                                <ComboBoxItem Content="44100" Tag="44100"/>
                                <ComboBoxItem Content="48000" Tag="48000" IsSelected="True"/>
                                <ComboBoxItem Content="96000" Tag="96000"/>
                            </ComboBox>
                            <TextBlock Text="Bit Depth" Style="{StaticResource Label}"/>
                            <ComboBox Name="WPFInputBitDepth" Margin="0,0,0,16">
                                <ComboBoxItem Content="16-bit"         Tag="16" IsSelected="True"/>
                                <ComboBoxItem Content="24-bit"         Tag="24"/>
                                <ComboBoxItem Content="32-bit (float)" Tag="32"/>
                            </ComboBox>
                            <Button Name="WPFApplyInputFormat" Style="{StaticResource BtnAccent}"
                                    Content="Apply Input Format" Margin="0,0,0,20" IsEnabled="False"/>
                            <Separator Background="#2D3561" Margin="0,0,0,16"/>
                            <TextBlock Text="Audio Enhancements" Style="{StaticResource SectionHeader}"/>
                            <TextBlock Text="Toggle system audio enhancements for this device."
                                       Style="{StaticResource Label}" TextWrapping="Wrap" Margin="0,0,0,10"/>
                            <ToggleButton Name="WPFInputEnhancementsToggle" Style="{StaticResource MuteBtn}"
                                          Content="Enhancements: ON" MinWidth="180" IsEnabled="False"
                                          HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 4: PROFILES -->
            <TabItem Name="WPFTab4">
                <Grid Margin="20">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Saved Profiles -->
                    <Border Grid.Column="0" Style="{StaticResource Card}">
                        <DockPanel>
                            <TextBlock DockPanel.Dock="Top" Text="Saved Profiles" Style="{StaticResource SectionHeader}"/>
                            <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="0,12,0,0">
                                <Button Name="WPFRestoreProfile" Style="{StaticResource BtnAccent}"
                                        Content="Restore" IsEnabled="False" Margin="0,0,8,0"/>
                                <Button Name="WPFDeleteProfile" Style="{StaticResource BtnBase}"
                                        Content="Delete" IsEnabled="False"/>
                            </StackPanel>
                            <TextBlock DockPanel.Dock="Bottom" Name="WPFProfileInfo" Text=""
                                       Foreground="#8892B0" FontSize="11" Margin="0,8,0,0" TextWrapping="Wrap"/>
                            <ListBox Name="WPFProfileList"/>
                        </DockPanel>
                    </Border>

                    <!-- Save Profile -->
                    <Border Grid.Column="2" Style="{StaticResource Card}">
                        <StackPanel>
                            <TextBlock Text="Save Current State" Style="{StaticResource SectionHeader}"/>
                            <TextBlock Text="Profile Name" Style="{StaticResource Label}"/>
                            <TextBox Name="WPFProfileNameInput" Margin="0,0,0,16"/>
                            <TextBlock Text="Capture" Style="{StaticResource Label}"/>
                            <CheckBox Name="WPFProfileSaveOutputDevice" Content="Default Output Device"
                                      IsChecked="True" Margin="0,4"/>
                            <CheckBox Name="WPFProfileSaveInputDevice"  Content="Default Input Device"
                                      IsChecked="True" Margin="0,4"/>
                            <CheckBox Name="WPFProfileSaveOutputVolume" Content="Output Volume and Mute"
                                      IsChecked="True" Margin="0,4"/>
                            <CheckBox Name="WPFProfileSaveInputVolume"  Content="Input Volume and Mute"
                                      IsChecked="True" Margin="0,4"/>
                            <CheckBox Name="WPFProfileSaveAppVolumes"   Content="Per-App Volumes"
                                      IsChecked="True" Margin="0,4,0,16"/>
                            <Button Name="WPFSaveProfile" Style="{StaticResource BtnAccent}"
                                    Content="Save Profile"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>

        </TabControl>
    </DockPanel>
</Window>

'@

# Parse XAML

$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' `
                       -replace 'xmlns:d="[^"]*"', '' `
                       -replace 'xmlns:mc="[^"]*"', '' `
                       -replace "x:Class=`"[^`"]*`"", ''

# Use Parse() so that WPF registers element names into the namescope.
# XmlNodeReader skips namescope registration, making FindName() return
# null for every control even though the Window loads and renders fine.
$sync.Form = [Windows.Markup.XamlReader]::Parse($inputXML)

# Bind all named controls into $sync by walking the logical tree
# (Parse() registers names, so FindName works correctly here)

[xml]$xaml = $inputXML
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    $ctrlName = $_.GetAttribute('Name')
    if ($ctrlName) {
        $ctrl = $sync.Form.FindName($ctrlName)
        if ($ctrl) { $sync[$ctrlName] = $ctrl }
    }
}

# Build runspace pool here (after all private/public functions are defined)
# so worker threads have every function available.
# Add-Type (CoreAudio) loads into the .NET AppDomain shared by all runspaces,
# so the C# types are automatically accessible without re-registering them.

$_iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$_iss.Variables.Add(
    [System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('sync', $sync, '')
)

# Only inject our own functions - avoids enumerating hundreds of built-in PS functions
$_amPattern = '^(Get-Audio|Set-Audio|Get-Device|Set-Device|Get-App|Set-App|' +
              'Set-Default|Remove-Audio|Save-Audio|Restore-Audio|' +
              'Initialize-|Update-|Invoke-WPF|Set-WPFStatus|Invoke-Audio)'
Get-Command -CommandType Function |
    Where-Object { $_.Name -match $_amPattern } |
    ForEach-Object {
        try {
            $entry = [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new(
                $_.Name, $_.ScriptBlock.ToString()
            )
            $_iss.Commands.Add($entry)
        } catch {}
    }

$sync.RunspacePool = [runspacefactory]::CreateRunspacePool(1, [Environment]::ProcessorCount, $_iss, $Host)
$sync.RunspacePool.Open()

# Version label

$sync.WPFVersionLabel.Text = "v$($sync.Version)"

# Tab navigation

foreach ($tabBtn in @('WPFTab1BT','WPFTab2BT','WPFTab3BT','WPFTab4BT')) {
    $btnName = $tabBtn
    $sync[$btnName].Add_Click({
        Invoke-WPFTab -ClickedTab $btnName
    }.GetNewClosure())
}

# Master volume slider

$sync.WPFMasterVolumeSlider.Add_ValueChanged({
    $pct = [math]::Round($sync.WPFMasterVolumeSlider.Value)
    $sync.WPFMasterVolumeLabel.Text = "$pct%"
})

$sync.WPFMasterVolumeSlider.Add_PreviewMouseLeftButtonUp({
    $level = $sync.WPFMasterVolumeSlider.Value / 100.0
    if ($sync.SelectedOutputId) {
        Invoke-AudioManagerRunspace { Set-DeviceVolume -DeviceId $sync.SelectedOutputId -Level $level }
    }
})

$sync.WPFMasterMuteButton.Add_Click({
    $muted = $sync.WPFMasterMuteButton.IsChecked
    if ($sync.SelectedOutputId) {
        Invoke-AudioManagerRunspace { Set-DeviceMute -DeviceId $sync.SelectedOutputId -Muted $muted }
    }
    $sync.WPFMasterMuteButton.Content = if ($muted) { "[X]" } else { "Mute" }
})

# Output device list selection

$sync.WPFOutputDeviceList.Add_SelectionChanged({
    $selected = $sync.WPFOutputDeviceList.SelectedItem
    if (-not $selected) {
        $sync.WPFSetDefaultOutput.IsEnabled   = $false
        $sync.WPFOutputVolumeSlider.IsEnabled = $false
        $sync.WPFOutputMuteButton.IsEnabled   = $false
        return
    }
    $sync.SelectedOutputId = $selected.Tag
    $sync.WPFSetDefaultOutput.IsEnabled   = $true
    $sync.WPFOutputVolumeSlider.IsEnabled = $true
    $sync.WPFOutputMuteButton.IsEnabled   = $true

    $dev = $sync.RenderDevices | Where-Object { $_.DeviceId -eq $selected.Tag } | Select-Object -First 1
    if ($dev) {
        $pct = [math]::Round($dev.VolumeScalar * 100)
        $sync.WPFOutputVolumeSlider.Value   = $pct
        $sync.WPFOutputVolumeLabel.Text     = "$pct%"
        $sync.WPFOutputMuteButton.IsChecked = $dev.IsMuted
        $sync.WPFMasterVolumeSlider.Value   = $pct
        $sync.WPFMasterVolumeLabel.Text     = "$pct%"
        $sync.WPFMasterMuteButton.IsChecked = $dev.IsMuted
        $sync.WPFMasterMuteButton.Content   = if ($dev.IsMuted) { "[X]" } else { "Mute" }
    }
})

$sync.WPFOutputVolumeSlider.Add_PreviewMouseLeftButtonUp({
    $level = $sync.WPFOutputVolumeSlider.Value / 100.0
    $pct   = [math]::Round($sync.WPFOutputVolumeSlider.Value)
    $sync.WPFOutputVolumeLabel.Text   = "$pct%"
    $sync.WPFMasterVolumeSlider.Value = $pct
    $sync.WPFMasterVolumeLabel.Text   = "$pct%"
    if ($sync.SelectedOutputId) {
        Invoke-AudioManagerRunspace { Set-DeviceVolume -DeviceId $sync.SelectedOutputId -Level $level }
    }
})

$sync.WPFOutputMuteButton.Add_Click({
    $muted = $sync.WPFOutputMuteButton.IsChecked
    if ($sync.SelectedOutputId) {
        Invoke-AudioManagerRunspace { Set-DeviceMute -DeviceId $sync.SelectedOutputId -Muted $muted }
    }
})

# Input device list selection

$sync.WPFInputDeviceList.Add_SelectionChanged({
    $selected = $sync.WPFInputDeviceList.SelectedItem
    if (-not $selected) {
        $sync.WPFSetDefaultInput.IsEnabled   = $false
        $sync.WPFInputVolumeSlider.IsEnabled = $false
        $sync.WPFInputMuteButton.IsEnabled   = $false
        return
    }
    $sync.SelectedInputId = $selected.Tag
    $sync.WPFSetDefaultInput.IsEnabled   = $true
    $sync.WPFInputVolumeSlider.IsEnabled = $true
    $sync.WPFInputMuteButton.IsEnabled   = $true

    $dev = $sync.CaptureDevices | Where-Object { $_.DeviceId -eq $selected.Tag } | Select-Object -First 1
    if ($dev) {
        $pct = [math]::Round($dev.VolumeScalar * 100)
        $sync.WPFInputVolumeSlider.Value   = $pct
        $sync.WPFInputVolumeLabel.Text     = "$pct%"
        $sync.WPFInputMuteButton.IsChecked = $dev.IsMuted
    }
})

$sync.WPFInputVolumeSlider.Add_PreviewMouseLeftButtonUp({
    $level = $sync.WPFInputVolumeSlider.Value / 100.0
    $pct   = [math]::Round($sync.WPFInputVolumeSlider.Value)
    $sync.WPFInputVolumeLabel.Text = "$pct%"
    if ($sync.SelectedInputId) {
        Invoke-AudioManagerRunspace { Set-DeviceVolume -DeviceId $sync.SelectedInputId -Level $level }
    }
})

$sync.WPFInputMuteButton.Add_Click({
    $muted = $sync.WPFInputMuteButton.IsChecked
    if ($sync.SelectedInputId) {
        Invoke-AudioManagerRunspace { Set-DeviceMute -DeviceId $sync.SelectedInputId -Muted $muted }
    }
})

# Formats tab device picker change

$sync.WPFFormatOutputDevice.Add_SelectionChanged({ Update-OutputFormatDisplay })
$sync.WPFFormatInputDevice.Add_SelectionChanged({  Update-InputFormatDisplay  })

# Enhancements toggles

$sync.WPFOutputEnhancementsToggle.Add_Click({
    $enabled  = $sync.WPFOutputEnhancementsToggle.IsChecked
    $selected = $sync.WPFFormatOutputDevice.SelectedItem
    if (-not $selected) { return }
    $deviceId = $selected.Tag
    Invoke-AudioManagerRunspace {
        $ok = Set-AudioEnhancement -DeviceId $deviceId -Enabled $enabled
        Invoke-WPFUIThread {
            $sync.WPFOutputEnhancementsToggle.Content = if ($enabled) { "Enhancements: ON" } else { "Enhancements: OFF" }
            Set-WPFStatus (if ($ok) { "Output enhancements updated." } else { "Failed to change enhancements." })
        }
    }
})

$sync.WPFInputEnhancementsToggle.Add_Click({
    $enabled  = $sync.WPFInputEnhancementsToggle.IsChecked
    $selected = $sync.WPFFormatInputDevice.SelectedItem
    if (-not $selected) { return }
    $deviceId = $selected.Tag
    Invoke-AudioManagerRunspace {
        $ok = Set-AudioEnhancement -DeviceId $deviceId -Enabled $enabled
        Invoke-WPFUIThread {
            $sync.WPFInputEnhancementsToggle.Content = if ($enabled) { "Enhancements: ON" } else { "Enhancements: OFF" }
            Set-WPFStatus (if ($ok) { "Input enhancements updated." } else { "Failed to change enhancements." })
        }
    }
})

# Profile list selection

$sync.WPFProfileList.Add_SelectionChanged({
    $selected   = $sync.WPFProfileList.SelectedItem
    $hasProfile = ($selected -and $selected.Tag)
    $sync.WPFRestoreProfile.IsEnabled = $hasProfile
    $sync.WPFDeleteProfile.IsEnabled  = $hasProfile
    if ($hasProfile) {
        $p    = $selected.Tag
        $info = "Created: $($p.created)"
        if ($p.defaultOutputDeviceName) { $info += "`nOutput: $($p.defaultOutputDeviceName)" }
        if ($p.defaultInputDeviceName)  { $info += "`nInput:  $($p.defaultInputDeviceName)" }
        $sync.WPFProfileInfo.Text = $info
    }
})

# Button dispatcher

foreach ($btnName in @(
    'WPFRefreshButton', 'WPFRefreshApps',
    'WPFSetDefaultOutput', 'WPFSetDefaultInput',
    'WPFApplyOutputFormat', 'WPFApplyInputFormat',
    'WPFSaveProfile', 'WPFRestoreProfile', 'WPFDeleteProfile'
)) {
    $name = $btnName
    if ($sync[$name]) {
        $sync[$name].Add_Click({
            Invoke-WPFButton -ClickedButton $name
        }.GetNewClosure())
    }
}

# Initial data load

Set-WPFStatus "Loading audio devices..."
Invoke-WPFRefreshAll

# Show window

$sync.Form.ShowDialog() | Out-Null

# Cleanup

$sync.RunspacePool.Close()
$sync.RunspacePool.Dispose()
Stop-Transcript -ErrorAction SilentlyContinue


