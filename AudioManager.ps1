#Requires -Version 5.1
# â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—
# â•‘           Audio Manager v26.04.20 â€” by PeterYama                â•‘
# â•‘   iwr -useb https://raw.githubusercontent.com/PeterYama/        â•‘
# â•‘       audio-manager/main/AudioManager.ps1 | iex                 â•‘
# â•‘   https://github.com/PeterYama/audio-manager                    â•‘
# â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#Requires -Version 5.1

$script:AMVersion = "26.04.20"

# ─── Elevation check ─────────────────────────────────────────────────────────

$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin   = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    if ($PSCommandPath) {
        $target = $PSCommandPath
    } else {
        # Script was IEX'd — save to temp and relaunch
        $target = "$env:TEMP\AudioManager_elevated.ps1"
        try {
            $scriptContent = (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/PeterYama/audio-manager/main/AudioManager.ps1" -UseBasicParsing)
            Set-Content -Path $target -Value $scriptContent -Encoding UTF8
        } catch {
            Write-Error "Could not download script for elevated relaunch. Run as Administrator manually."
            exit 1
        }
    }
    $args = "-ExecutionPolicy Bypass -NonInteractive -File `"$target`""
    if (Get-Command wt.exe -ErrorAction SilentlyContinue) {
        Start-Process wt.exe -Verb RunAs -ArgumentList "powershell.exe $args"
    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList $args
    }
    exit
}

# ─── Assembly loading ─────────────────────────────────────────────────────────

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ─── Logging ─────────────────────────────────────────────────────────────────

$logDir = "$env:LOCALAPPDATA\AudioManager\logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
Start-Transcript -Path "$logDir\AudioManager_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" -Append -ErrorAction SilentlyContinue

# ─── Profiles directory ───────────────────────────────────────────────────────

$profilesDir = "$env:APPDATA\AudioManager"
if (-not (Test-Path $profilesDir)) { New-Item -ItemType Directory -Path $profilesDir -Force | Out-Null }

# ─── Shared synchronized hashtable ───────────────────────────────────────────

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

# ─── Runspace pool (shared across background jobs) ────────────────────────────

$sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$sessionState.Variables.Add(
    [System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('sync', $sync, '')
)
$sync.RunspacePool = [runspacefactory]::CreateRunspacePool(1, [Environment]::ProcessorCount, $sessionState, $Host)
$sync.RunspacePool.Open()

$script:CoreAudioCSharp = @'
using System;
using System.Runtime.InteropServices;

namespace AudioManager
{
    // ─── Enumerations ────────────────────────────────────────────────────────────

    public enum EDataFlow { eRender = 0, eCapture = 1, eAll = 2 }
    public enum ERole    { eConsole = 0, eMultimedia = 1, eCommunications = 2 }
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

    // ─── Structs ─────────────────────────────────────────────────────────────────

    [StructLayout(LayoutKind.Sequential)]
    public struct WAVEFORMATEX
    {
        public ushort wFormatTag;
        public ushort nChannels;
        public uint   nSamplesPerSec;
        public uint   nAvgBytesPerSec;
        public ushort nBlockAlign;
        public ushort wBitsPerSample;
        public ushort cbSize;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct WAVEFORMATEXTENSIBLE
    {
        [FieldOffset(0)]  public WAVEFORMATEX Format;
        [FieldOffset(18)] public ushort       wValidBitsPerSample;
        [FieldOffset(20)] public uint         dwChannelMask;
        [FieldOffset(24)] public Guid         SubFormat;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PropertyKey
    {
        public Guid  fmtid;
        public uint  pid;
    }

    // ─── IMMDeviceEnumerator ─────────────────────────────────────────────────────

    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceEnumerator
    {
        int EnumAudioEndpoints(
            [In] EDataFlow dataFlow,
            [In] DeviceState stateMask,
            [Out] out IMMDeviceCollection devices);

        int GetDefaultAudioEndpoint(
            [In] EDataFlow dataFlow,
            [In] ERole role,
            [Out] out IMMDevice endpoint);

        int GetDevice(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pwstrId,
            [Out] out IMMDevice device);

        int RegisterEndpointNotificationCallback(IntPtr client);
        int UnregisterEndpointNotificationCallback(IntPtr client);
    }

    // ─── IMMDeviceCollection ─────────────────────────────────────────────────────

    [Guid("0BD7A1BE-7A1A-44DB-8397-BE5155E7F6E1")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceCollection
    {
        int GetCount([Out] out uint count);
        int Item([In] uint index, [Out] out IMMDevice device);
    }

    // ─── IMMDevice ───────────────────────────────────────────────────────────────

    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice
    {
        int Activate(
            [In] ref Guid iid,
            [In] int dwClsCtx,
            [In] IntPtr pActivationParams,
            [Out, MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);

        int OpenPropertyStore(
            [In] int stgmAccess,
            [Out] out IPropertyStore properties);

        int GetId([Out, MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetState([Out] out DeviceState state);
    }

    // ─── IPropertyStore ──────────────────────────────────────────────────────────

    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore
    {
        int GetCount([Out] out int count);
        int GetAt([In] int index, [Out] out PropertyKey key);
        int GetValue([In] ref PropertyKey key, [Out] out PropVariant value);
        int SetValue([In] ref PropertyKey key, [In] ref PropVariant value);
        int Commit();
    }

    // ─── PropVariant (simplified) ────────────────────────────────────────────────

    [StructLayout(LayoutKind.Sequential)]
    public struct PropVariant
    {
        public ushort vt;
        public ushort reserved1;
        public ushort reserved2;
        public ushort reserved3;
        public IntPtr  data;
        public IntPtr  data2;

        public string GetStringValue()
        {
            if (vt == 31) // VT_LPWSTR
                return Marshal.PtrToStringUni(data);
            return string.Empty;
        }
    }

    // ─── IAudioEndpointVolume ────────────────────────────────────────────────────

    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioEndpointVolume
    {
        int RegisterControlChangeNotify(IntPtr callback);
        int UnregisterControlChangeNotify(IntPtr callback);
        int GetChannelCount([Out] out int count);
        int SetMasterVolumeLevel([In] float level, [In] ref Guid eventContext);
        int SetMasterVolumeLevelScalar([In] float level, [In] ref Guid eventContext);
        int GetMasterVolumeLevel([Out] out float level);
        int GetMasterVolumeLevelScalar([Out] out float level);
        int SetChannelVolumeLevel([In] uint channel, [In] float level, [In] ref Guid eventContext);
        int SetChannelVolumeLevelScalar([In] uint channel, [In] float level, [In] ref Guid eventContext);
        int GetChannelVolumeLevel([In] uint channel, [Out] out float level);
        int GetChannelVolumeLevelScalar([In] uint channel, [Out] out float level);
        int SetMute([In, MarshalAs(UnmanagedType.Bool)] bool isMuted, [In] ref Guid eventContext);
        int GetMute([Out, MarshalAs(UnmanagedType.Bool)] out bool isMuted);
        int GetVolumeStepInfo([Out] out uint step, [Out] out uint stepCount);
        int VolumeStepUp([In] ref Guid eventContext);
        int VolumeStepDown([In] ref Guid eventContext);
        int QueryHardwareSupport([Out] out uint hardwareSupportMask);
        int GetVolumeRange([Out] out float volumeMin, [Out] out float volumeMax, [Out] out float volumeIncrement);
    }

    // ─── IAudioSessionManager2 ───────────────────────────────────────────────────

    [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioSessionManager2
    {
        int GetAudioSessionControl(
            [In] ref Guid audioSessionGuid,
            [In] uint streamFlags,
            [Out] out IAudioSessionControl session);

        int GetSimpleAudioVolume(
            [In] ref Guid audioSessionGuid,
            [In] uint streamFlags,
            [Out] out ISimpleAudioVolume audioVolume);

        int GetSessionEnumerator([Out] out IAudioSessionEnumerator sessionEnum);
        int RegisterSessionNotification(IntPtr notification);
        int UnregisterSessionNotification(IntPtr notification);
        int RegisterDuckNotification([MarshalAs(UnmanagedType.LPWStr)] string sessionId, IntPtr duckNotification);
        int UnregisterDuckNotification(IntPtr duckNotification);
    }

    // ─── IAudioSessionEnumerator ─────────────────────────────────────────────────

    [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioSessionEnumerator
    {
        int GetCount([Out] out int sessionCount);
        int GetSession([In] int sessionIndex, [Out] out IAudioSessionControl session);
    }

    // ─── IAudioSessionControl ────────────────────────────────────────────────────

    [Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioSessionControl
    {
        int GetState([Out] out AudioSessionState state);
        int GetDisplayName([Out, MarshalAs(UnmanagedType.LPWStr)] out string displayName);
        int SetDisplayName([In, MarshalAs(UnmanagedType.LPWStr)] string value, [In] ref Guid eventContext);
        int GetIconPath([Out, MarshalAs(UnmanagedType.LPWStr)] out string iconPath);
        int SetIconPath([In, MarshalAs(UnmanagedType.LPWStr)] string value, [In] ref Guid eventContext);
        int GetGroupingParam([Out] out Guid groupingParam);
        int SetGroupingParam([In] ref Guid Override, [In] ref Guid eventContext);
        int RegisterAudioSessionNotification(IntPtr NewNotifications);
        int UnregisterAudioSessionNotification(IntPtr NewNotifications);
    }

    // ─── IAudioSessionControl2 ───────────────────────────────────────────────────

    [Guid("BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioSessionControl2
    {
        // Inherited from IAudioSessionControl
        int GetState([Out] out AudioSessionState state);
        int GetDisplayName([Out, MarshalAs(UnmanagedType.LPWStr)] out string displayName);
        int SetDisplayName([In, MarshalAs(UnmanagedType.LPWStr)] string value, [In] ref Guid eventContext);
        int GetIconPath([Out, MarshalAs(UnmanagedType.LPWStr)] out string iconPath);
        int SetIconPath([In, MarshalAs(UnmanagedType.LPWStr)] string value, [In] ref Guid eventContext);
        int GetGroupingParam([Out] out Guid groupingParam);
        int SetGroupingParam([In] ref Guid Override, [In] ref Guid eventContext);
        int RegisterAudioSessionNotification(IntPtr NewNotifications);
        int UnregisterAudioSessionNotification(IntPtr NewNotifications);
        // IAudioSessionControl2-specific
        int GetSessionIdentifier([Out, MarshalAs(UnmanagedType.LPWStr)] out string retVal);
        int GetSessionInstanceIdentifier([Out, MarshalAs(UnmanagedType.LPWStr)] out string retVal);
        int GetProcessId([Out] out uint retVal);
        int IsSystemSoundsSession();
        int SetDuckingPreference([In, MarshalAs(UnmanagedType.Bool)] bool optOut);
    }

    // ─── ISimpleAudioVolume ──────────────────────────────────────────────────────

    [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISimpleAudioVolume
    {
        int SetMasterVolume([In] float fLevel, [In] ref Guid EventContext);
        int GetMasterVolume([Out] out float pfLevel);
        int SetMute([In, MarshalAs(UnmanagedType.Bool)] bool bMute, [In] ref Guid EventContext);
        int GetMute([Out, MarshalAs(UnmanagedType.Bool)] out bool pbMute);
    }

    // ─── IPolicyConfig (undocumented, stable since Vista) ────────────────────────

    [Guid("F8679F50-850A-41CF-9C72-430F290290C8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPolicyConfig
    {
        int GetMixFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [Out] out IntPtr ppFormat);
        int GetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In, MarshalAs(UnmanagedType.Bool)] bool bDefault, [Out] out IntPtr ppFormat);
        int ResetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName);
        int SetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In] ref WAVEFORMATEX pEndpointFormat, IntPtr MixFormat);
        int GetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In, MarshalAs(UnmanagedType.Bool)] bool bDefault, [Out] out long pmftDefaultPeriod, [Out] out long pmftMinimumPeriod);
        int SetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In] ref long pmftPeriod);
        int GetShareMode([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [Out] out DeviceShareMode pMode);
        int SetShareMode([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In] DeviceShareMode mode);
        int GetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In, MarshalAs(UnmanagedType.Bool)] bool bFxStore, [In] ref PropertyKey key, [Out] out PropVariant pv);
        int SetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In, MarshalAs(UnmanagedType.Bool)] bool bFxStore, [In] ref PropertyKey key, [In] ref PropVariant pv);
        int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In] ERole role);
        int SetEndpointVisibility([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In, MarshalAs(UnmanagedType.Bool)] bool bVisible);
    }

    public enum DeviceShareMode { Shared = 0, Exclusive = 1 }

    // ─── COM CoCreate helpers ────────────────────────────────────────────────────

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    public class MMDeviceEnumeratorComObject { }

    [ComImport, Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
    public class CPolicyConfigClient { }

    // ─── Static helper ───────────────────────────────────────────────────────────

    public static class AudioManagerHelper
    {
        private static readonly Guid DEVINTERFACE_AUDIO_RENDER  = new Guid("E6327CAD-DCEC-4949-AE8A-991E976A79D2");
        private static readonly Guid DEVINTERFACE_AUDIO_CAPTURE = new Guid("2EEF81BE-33FA-4800-9670-1CD474972C3F");

        public static IMMDeviceEnumerator CreateDeviceEnumerator()
        {
            return (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
        }

        public static IPolicyConfig CreatePolicyConfig()
        {
            return (IPolicyConfig)new CPolicyConfigClient();
        }

        public static PropertyKey PKEY_Device_FriendlyName()
        {
            return new PropertyKey
            {
                fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"),
                pid   = 14
            };
        }

        public static PropertyKey PKEY_AudioEndpoint_Disable_SysFx()
        {
            return new PropertyKey
            {
                fmtid = new Guid("1da5d803-d492-4edd-8c23-e0c0ffee7f0e"),
                pid   = 5
            };
        }

        public static PropertyKey PKEY_AudioEngine_DeviceFormat()
        {
            return new PropertyKey
            {
                fmtid = new Guid("f19f064d-082c-4e27-bc73-6882a1bb8e4c"),
                pid   = 0
            };
        }

        public static Guid IID_IAudioEndpointVolume   = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
        public static Guid IID_IAudioSessionManager2  = new Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F");
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
        $enumerator = [AudioManager.AudioManagerHelper]::CreateDeviceEnumerator()
        foreach ($flow in @([AudioManager.EDataFlow]::eRender, [AudioManager.EDataFlow]::eCapture)) {
            $collection = $null
            $enumerator.EnumAudioEndpoints($flow, [AudioManager.DeviceState]::Active, [ref]$collection) | Out-Null
            $count = 0
            $collection.GetCount([ref]$count) | Out-Null
            $devices = @()
            for ($i = 0; $i -lt $count; $i++) {
                $device = $null
                $collection.Item($i, [ref]$device) | Out-Null

                $id = ""
                $device.GetId([ref]$id) | Out-Null

                $store = $null
                $device.OpenPropertyStore(0, [ref]$store) | Out-Null  # STGM_READ = 0
                $key = [AudioManager.AudioManagerHelper]::PKEY_Device_FriendlyName()
                $pv  = New-Object AudioManager.PropVariant
                $store.GetValue([ref]$key, [ref]$pv) | Out-Null
                $name = $pv.GetStringValue()
                if ([string]::IsNullOrEmpty($name)) { $name = "Unknown Device" }

                # Get endpoint volume interface
                $volIid = [AudioManager.AudioManagerHelper]::IID_IAudioEndpointVolume
                $volObj = $null
                $device.Activate([ref]$volIid, 23, [IntPtr]::Zero, [ref]$volObj) | Out-Null
                $vol = [AudioManager.IAudioEndpointVolume]$volObj

                $level = 0.0
                $vol.GetMasterVolumeLevelScalar([ref]$level) | Out-Null
                $muted = $false
                $vol.GetMute([ref]$muted) | Out-Null

                $devices += [PSCustomObject]@{
                    DeviceId      = $id
                    Name          = $name
                    Flow          = $flow
                    VolumeScalar  = $level
                    IsMuted       = $muted
                    VolumeInterface = $vol
                }
            }
            if ($flow -eq [AudioManager.EDataFlow]::eRender) {
                $result.Render = $devices
            } else {
                $result.Capture = $devices
            }
        }
    } catch {
        Write-Warning "Get-AudioDevices error: $_"
    }
    return $result
}


function Get-AudioEnhancement {
    param([Parameter(Mandatory)][string]$DeviceId)
    # DeviceId format: {0.0.0.00000000}.{GUID}  — extract the GUID portion
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
    $sessions = @()
    try {
        $enumerator = [AudioManager.AudioManagerHelper]::CreateDeviceEnumerator()
        $defaultDev = $null
        $enumerator.GetDefaultAudioEndpoint(
            [AudioManager.EDataFlow]::eRender,
            [AudioManager.ERole]::eConsole,
            [ref]$defaultDev
        ) | Out-Null

        $mgr2Iid = [AudioManager.AudioManagerHelper]::IID_IAudioSessionManager2
        $mgr2Obj = $null
        $defaultDev.Activate([ref]$mgr2Iid, 23, [IntPtr]::Zero, [ref]$mgr2Obj) | Out-Null
        $mgr2 = [AudioManager.IAudioSessionManager2]$mgr2Obj

        $sessionEnum = $null
        $mgr2.GetSessionEnumerator([ref]$sessionEnum) | Out-Null

        $count = 0
        $sessionEnum.GetCount([ref]$count) | Out-Null

        for ($i = 0; $i -lt $count; $i++) {
            $ctrl = $null
            $sessionEnum.GetSession($i, [ref]$ctrl) | Out-Null

            try {
                $ctrl2    = [AudioManager.IAudioSessionControl2]$ctrl
                $simpleVol = [AudioManager.ISimpleAudioVolume]$ctrl

                $pid = 0
                $ctrl2.GetProcessId([ref]$pid) | Out-Null

                # Skip the system sounds session (PID 0)
                if ($pid -eq 0) { continue }

                $displayName = ""
                $ctrl2.GetDisplayName([ref]$displayName) | Out-Null

                $proc = $null
                if ([string]::IsNullOrEmpty($displayName)) {
                    $proc = Get-Process -Id $pid -ErrorAction SilentlyContinue
                    $displayName = if ($proc -and $proc.MainWindowTitle) {
                        $proc.MainWindowTitle
                    } elseif ($proc) {
                        $proc.ProcessName
                    } else {
                        "PID $pid"
                    }
                }

                $vol   = 0.0
                $muted = $false
                $simpleVol.GetMasterVolume([ref]$vol)   | Out-Null
                $simpleVol.GetMute([ref]$muted)          | Out-Null

                $state = [AudioManager.AudioSessionState]::Inactive
                $ctrl.GetState([ref]$state) | Out-Null

                $sessions += [PSCustomObject]@{
                    SessionKey    = "$pid-$i"
                    Name          = $displayName
                    ProcessId     = $pid
                    PidLabel      = "PID: $pid"
                    VolumePercent = [math]::Round($vol * 100)
                    VolumeLabel   = "$([math]::Round($vol * 100))%"
                    IsMuted       = $muted
                    State         = $state
                    Icon          = Get-ProcessIcon -ProcessId $pid
                    SimpleVolume  = $simpleVol
                    SessionControl = $ctrl
                }
            } catch {
                # Session may have expired, skip it
            }
        }
    } catch {
        Write-Warning "Get-AudioSessions error: $_"
    }
    return $sessions
}

function Get-ProcessIcon {
    param([uint32]$ProcessId)
    $map = @{
        'chrome'         = '🌐'
        'firefox'        = '🦊'
        'msedge'         = '🌐'
        'spotify'        = '🎵'
        'discord'        = '💬'
        'slack'          = '💼'
        'teams'          = '👥'
        'zoom'           = '📹'
        'vlc'            = '🎬'
        'obs64'          = '🎥'
        'obs32'          = '🎥'
        'steam'          = '🎮'
        'epicgameslauncher' = '🎮'
        'mpc-hc64'       = '🎬'
        'foobar2000'     = '🎶'
        'itunes'         = '🎵'
        'winamp'         = '🎵'
    }
    try {
        $proc = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        if ($proc) {
            $key = $proc.ProcessName.ToLower()
            if ($map.ContainsKey($key)) { return $map[$key] }
        }
    } catch {}
    return '🔊'
}


function Get-DeviceFormat {
    param([Parameter(Mandatory)][string]$DeviceId)
    try {
        $policy    = [AudioManager.AudioManagerHelper]::CreatePolicyConfig()
        $fmtPtr    = [IntPtr]::Zero
        $policy.GetDeviceFormat($DeviceId, $false, [ref]$fmtPtr) | Out-Null

        if ($fmtPtr -eq [IntPtr]::Zero) { return $null }

        $fmt = [System.Runtime.InteropServices.Marshal]::PtrToStructure(
            $fmtPtr,
            [AudioManager.WAVEFORMATEX]
        )

        return [PSCustomObject]@{
            SampleRate  = $fmt.nSamplesPerSec
            BitDepth    = $fmt.wBitsPerSample
            Channels    = $fmt.nChannels
            FormatTag   = $fmt.wFormatTag
            Description = "$($fmt.nSamplesPerSec) Hz / $($fmt.wBitsPerSample)-bit / $($fmt.nChannels)ch"
        }
    } catch {
        Write-Warning "Get-DeviceFormat error: $_"
        return $null
    }
}


function Get-DeviceMute {
    param([Parameter(Mandatory)][string]$DeviceId)
    try {
        $enumerator = [AudioManager.AudioManagerHelper]::CreateDeviceEnumerator()
        $device     = $null
        $enumerator.GetDevice($DeviceId, [ref]$device) | Out-Null
        $iid    = [AudioManager.AudioManagerHelper]::IID_IAudioEndpointVolume
        $volObj = $null
        $device.Activate([ref]$iid, 23, [IntPtr]::Zero, [ref]$volObj) | Out-Null
        $vol    = [AudioManager.IAudioEndpointVolume]$volObj
        $muted  = $false
        $vol.GetMute([ref]$muted) | Out-Null
        return $muted
    } catch {
        Write-Warning "Get-DeviceMute error: $_"
        return $false
    }
}


function Get-DeviceVolume {
    param([Parameter(Mandatory)][string]$DeviceId)
    try {
        $enumerator = [AudioManager.AudioManagerHelper]::CreateDeviceEnumerator()
        $device     = $null
        $enumerator.GetDevice($DeviceId, [ref]$device) | Out-Null
        $iid    = [AudioManager.AudioManagerHelper]::IID_IAudioEndpointVolume
        $volObj = $null
        $device.Activate([ref]$iid, 23, [IntPtr]::Zero, [ref]$volObj) | Out-Null
        $vol    = [AudioManager.IAudioEndpointVolume]$volObj
        $level  = 0.0
        $vol.GetMasterVolumeLevelScalar([ref]$level) | Out-Null
        return $level
    } catch {
        Write-Warning "Get-DeviceVolume error: $_"
        return 0.0
    }
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
        [Parameter(Mandatory)][string]$SessionKey,
        [Parameter(Mandatory)][bool]$Muted
    )
    try {
        $session = $sync.AudioSessions | Where-Object { $_.SessionKey -eq $SessionKey } | Select-Object -First 1
        if ($session -and $session.SimpleVolume) {
            $guid = [guid]::Empty
            $session.SimpleVolume.SetMute($Muted, [ref]$guid) | Out-Null
            $session.IsMuted = $Muted
        }
    } catch {
        Write-Warning "Set-AppMute error: $_"
    }
}


function Set-AppVolume {
    param(
        [Parameter(Mandatory)][string]$SessionKey,   # "pid-index"
        [Parameter(Mandatory)][float]$Level           # 0.0 – 1.0
    )
    try {
        $Level = [math]::Max(0.0, [math]::Min(1.0, $Level))
        $session = $sync.AudioSessions | Where-Object { $_.SessionKey -eq $SessionKey } | Select-Object -First 1
        if ($session -and $session.SimpleVolume) {
            $guid = [guid]::Empty
            $session.SimpleVolume.SetMasterVolume($Level, [ref]$guid) | Out-Null
            $session.VolumePercent = [math]::Round($Level * 100)
            $session.VolumeLabel   = "$($session.VolumePercent)%"
        }
    } catch {
        Write-Warning "Set-AppVolume error: $_"
    }
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
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][AudioManager.EDataFlow]$Flow
    )
    try {
        $policy = [AudioManager.AudioManagerHelper]::CreatePolicyConfig()
        # Set for all three roles
        foreach ($role in @(
            [AudioManager.ERole]::eConsole,
            [AudioManager.ERole]::eMultimedia,
            [AudioManager.ERole]::eCommunications
        )) {
            $policy.SetDefaultEndpoint($DeviceId, $role) | Out-Null
        }
        return $true
    } catch {
        Write-Warning "Set-DefaultAudioDevice error: $_"
        return $false
    }
}


function Set-DeviceFormat {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][uint32]$SampleRate,
        [Parameter(Mandatory)][uint16]$BitDepth
    )
    try {
        $fmt = New-Object AudioManager.WAVEFORMATEX
        # For PCM (16/24-bit) use format tag 1; for 32-bit float use 3
        $fmt.wFormatTag      = if ($BitDepth -eq 32) { [uint16]3 } else { [uint16]1 }
        $fmt.nChannels       = [uint16]2
        $fmt.nSamplesPerSec  = $SampleRate
        $fmt.wBitsPerSample  = $BitDepth
        $fmt.nBlockAlign     = [uint16](($fmt.nChannels * $fmt.wBitsPerSample) / 8)
        $fmt.nAvgBytesPerSec = $fmt.nSamplesPerSec * $fmt.nBlockAlign
        $fmt.cbSize          = [uint16]0

        $policy = [AudioManager.AudioManagerHelper]::CreatePolicyConfig()
        $policy.SetDeviceFormat($DeviceId, [ref]$fmt, [IntPtr]::Zero) | Out-Null
        return $true
    } catch {
        Write-Warning "Set-DeviceFormat error: $_"
        return $false
    }
}


function Set-DeviceMute {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][bool]$Muted
    )
    try {
        $enumerator = [AudioManager.AudioManagerHelper]::CreateDeviceEnumerator()
        $device     = $null
        $enumerator.GetDevice($DeviceId, [ref]$device) | Out-Null
        $iid    = [AudioManager.AudioManagerHelper]::IID_IAudioEndpointVolume
        $volObj = $null
        $device.Activate([ref]$iid, 23, [IntPtr]::Zero, [ref]$volObj) | Out-Null
        $vol    = [AudioManager.IAudioEndpointVolume]$volObj
        $guid   = [guid]::Empty
        $vol.SetMute($Muted, [ref]$guid) | Out-Null
        return $true
    } catch {
        Write-Warning "Set-DeviceMute error: $_"
        return $false
    }
}


function Set-DeviceVolume {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][float]$Level     # 0.0 – 1.0
    )
    try {
        $Level = [math]::Max(0.0, [math]::Min(1.0, $Level))
        $enumerator = [AudioManager.AudioManagerHelper]::CreateDeviceEnumerator()
        $device     = $null
        $enumerator.GetDevice($DeviceId, [ref]$device) | Out-Null
        $iid    = [AudioManager.AudioManagerHelper]::IID_IAudioEndpointVolume
        $volObj = $null
        $device.Activate([ref]$iid, 23, [IntPtr]::Zero, [ref]$volObj) | Out-Null
        $vol    = [AudioManager.IAudioEndpointVolume]$volObj
        $guid   = [guid]::Empty
        $vol.SetMasterVolumeLevelScalar($Level, [ref]$guid) | Out-Null
        return $true
    } catch {
        Write-Warning "Set-DeviceVolume error: $_"
        return $false
    }
}


function Initialize-ApplicationsTab {
    $sync.WPFAppSessionList.Items.Clear()

    if ($sync.AudioSessions.Count -eq 0) {
        $sync.WPFAppCount.Text = "No active audio sessions found."
        return
    }

    $sync.WPFAppCount.Text = "$($sync.AudioSessions.Count) session(s)"

    foreach ($session in ($sync.AudioSessions | Sort-Object Name)) {
        $sync.WPFAppSessionList.Items.Add($session) | Out-Null
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
            Set-WPFStatus "Setting default output to '$($selected.Content)'..."
            Invoke-AudioManagerRunspace {
                $ok = Set-DefaultAudioDevice -DeviceId $deviceId -Flow ([AudioManager.EDataFlow]::eRender)
                Invoke-WPFUIThread {
                    Set-WPFStatus (if ($ok) { "Default output set to '$($selected.Content)'." } else { "Failed to set default output." })
                }
            }
        }

        "WPFSetDefaultInput" {
            $selected = $sync.WPFInputDeviceList.SelectedItem
            if (-not $selected) { Set-WPFStatus "Select an input device first."; return }
            $deviceId = $selected.Tag
            Set-WPFStatus "Setting default input to '$($selected.Content)'..."
            Invoke-AudioManagerRunspace {
                $ok = Set-DefaultAudioDevice -DeviceId $deviceId -Flow ([AudioManager.EDataFlow]::eCapture)
                Invoke-WPFUIThread {
                    Set-WPFStatus (if ($ok) { "Default input set to '$($selected.Content)'." } else { "Failed to set default input." })
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
                Set-WPFStatus "Ready — $($sync.RenderDevices.Count) output, $($sync.CaptureDevices.Count) input device(s) found."
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
        <!-- Colors -->
        <SolidColorBrush x:Key="BgDark"     Color="#1A1A2E"/>
        <SolidColorBrush x:Key="BgMid"      Color="#16213E"/>
        <SolidColorBrush x:Key="BgLight"    Color="#0F3460"/>
        <SolidColorBrush x:Key="Accent"     Color="#E94560"/>
        <SolidColorBrush x:Key="AccentHover"Color="#FF6B81"/>
        <SolidColorBrush x:Key="TextPrimary"Color="#E0E0E0"/>
        <SolidColorBrush x:Key="TextMuted"  Color="#8892B0"/>
        <SolidColorBrush x:Key="Green"      Color="#43D97B"/>
        <SolidColorBrush x:Key="CardBg"     Color="#16213E"/>

        <!-- Base button style -->
        <Style x:Key="BtnBase" TargetType="Button">
            <Setter Property="Background"   Value="#0F3460"/>
            <Setter Property="Foreground"   Value="#E0E0E0"/>
            <Setter Property="BorderBrush"  Value="#E94560"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"      Value="14,7"/>
            <Setter Property="Cursor"       Value="Hand"/>
            <Setter Property="FontSize"     Value="12"/>
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
            <Setter Property="Background" Value="#E94560"/>
            <Setter Property="BorderBrush" Value="#E94560"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FF6B81"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Tab navigation toggle button -->
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
                                <Setter Property="Foreground"   Value="#E94560"/>
                                <Setter Property="BorderBrush"  Value="#E94560"/>
                                <Setter Property="FontWeight"   Value="SemiBold"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="#E0E0E0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Slider style -->
        <Style x:Key="AudioSlider" TargetType="Slider">
            <Setter Property="Minimum"          Value="0"/>
            <Setter Property="Maximum"          Value="100"/>
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
                            <Track x:Name="PART_Track" Grid.Row="1">
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

        <!-- Mute / toggle button style -->
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

        <!-- Card border style -->
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background"       Value="#16213E"/>
            <Setter Property="BorderBrush"      Value="#2D3561"/>
            <Setter Property="BorderThickness"  Value="1"/>
            <Setter Property="CornerRadius"     Value="8"/>
            <Setter Property="Padding"          Value="16"/>
        </Style>

        <!-- ListBox style -->
        <Style TargetType="ListBox">
            <Setter Property="Background"      Value="#0F1729"/>
            <Setter Property="Foreground"      Value="#E0E0E0"/>
            <Setter Property="BorderBrush"     Value="#2D3561"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="Padding"  Value="10,8"/>
            <Setter Property="Cursor"   Value="Hand"/>
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

        <!-- ComboBox style -->
        <Style TargetType="ComboBox">
            <Setter Property="Background"      Value="#0F3460"/>
            <Setter Property="Foreground"      Value="#E0E0E0"/>
            <Setter Property="BorderBrush"     Value="#2D3561"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="8,5"/>
        </Style>

        <!-- TextBox style -->
        <Style TargetType="TextBox">
            <Setter Property="Background"      Value="#0F1729"/>
            <Setter Property="Foreground"      Value="#E0E0E0"/>
            <Setter Property="BorderBrush"     Value="#2D3561"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="8,6"/>
            <Setter Property="CaretBrush"      Value="#E94560"/>
        </Style>

        <!-- Section header text -->
        <Style x:Key="SectionHeader" TargetType="TextBlock">
            <Setter Property="FontSize"   Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#E94560"/>
            <Setter Property="Margin"     Value="0,0,0,10"/>
        </Style>

        <!-- Label text -->
        <Style x:Key="Label" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#8892B0"/>
            <Setter Property="FontSize"   Value="11"/>
            <Setter Property="Margin"     Value="0,0,0,4"/>
        </Style>

        <!-- ScrollBar style -->
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="#16213E"/>
            <Setter Property="Width"      Value="6"/>
        </Style>
    </Window.Resources>

    <DockPanel>

        <!-- ═══════════════════════════════════════════════════════════════════
             HEADER BAR
        ════════════════════════════════════════════════════════════════════ -->
        <Border DockPanel.Dock="Top" Background="#16213E" BorderBrush="#2D3561"
                BorderThickness="0,0,0,1" Padding="20,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Title -->
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="🔊" FontSize="20" Margin="0,0,8,0" VerticalAlignment="Center"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="Audio Manager" FontSize="16" FontWeight="Bold" Foreground="#E0E0E0"/>
                        <TextBlock x:Name="WPFVersionLabel" Text="v0.0.0" FontSize="10" Foreground="#8892B0"/>
                    </StackPanel>
                </StackPanel>

                <!-- Master Volume -->
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Center"
                            VerticalAlignment="Center" Margin="20,0">
                    <TextBlock Text="MASTER" FontSize="10" FontWeight="SemiBold" Foreground="#8892B0"
                               VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <ToggleButton x:Name="WPFMasterMuteButton" Style="{StaticResource MuteBtn}"
                                  Content="🔇" FontSize="14" Width="36" Margin="0,0,10,0"
                                  ToolTip="Mute/Unmute master output"/>
                    <Slider x:Name="WPFMasterVolumeSlider" Style="{StaticResource AudioSlider}"
                            Width="200" VerticalAlignment="Center"/>
                    <TextBlock x:Name="WPFMasterVolumeLabel" Text="100%" Foreground="#E0E0E0"
                               FontWeight="SemiBold" Width="40" Margin="10,0,0,0" VerticalAlignment="Center"/>
                </StackPanel>

                <!-- Refresh -->
                <Button x:Name="WPFRefreshButton" Grid.Column="2" Style="{StaticResource BtnBase}"
                        Content="↺  Refresh" VerticalAlignment="Center" FontSize="12"/>
            </Grid>
        </Border>

        <!-- ═══════════════════════════════════════════════════════════════════
             TAB NAVIGATION
        ════════════════════════════════════════════════════════════════════ -->
        <Border DockPanel.Dock="Top" Background="#16213E" BorderBrush="#2D3561"
                BorderThickness="0,0,0,1">
            <StackPanel Orientation="Horizontal">
                <ToggleButton x:Name="WPFTab1BT" Style="{StaticResource TabBtn}" Content="🖥  Devices"     IsChecked="True"/>
                <ToggleButton x:Name="WPFTab2BT" Style="{StaticResource TabBtn}" Content="🎵  Applications"/>
                <ToggleButton x:Name="WPFTab3BT" Style="{StaticResource TabBtn}" Content="⚙  Formats"/>
                <ToggleButton x:Name="WPFTab4BT" Style="{StaticResource TabBtn}" Content="💾  Profiles"/>
            </StackPanel>
        </Border>

        <!-- ═══════════════════════════════════════════════════════════════════
             STATUS BAR
        ════════════════════════════════════════════════════════════════════ -->
        <Border DockPanel.Dock="Bottom" Background="#16213E" BorderBrush="#2D3561"
                BorderThickness="0,1,0,0" Padding="16,6">
            <TextBlock x:Name="WPFStatusBar" Text="Ready" Foreground="#8892B0" FontSize="11"/>
        </Border>

        <!-- ═══════════════════════════════════════════════════════════════════
             TAB CONTENT
        ════════════════════════════════════════════════════════════════════ -->
        <TabControl x:Name="WPFTabControl" Background="Transparent" BorderThickness="0">
            <TabControl.Resources>
                <Style TargetType="TabItem">
                    <Setter Property="Visibility" Value="Collapsed"/>
                </Style>
            </TabControl.Resources>

            <!-- ── TAB 1: DEVICES ────────────────────────────────────────── -->
            <TabItem x:Name="WPFTab1" IsSelected="True">
                <Grid Margin="20" VerticalAlignment="Stretch">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- OUTPUT -->
                    <Border Grid.Column="0" Style="{StaticResource Card}">
                        <DockPanel>
                            <TextBlock DockPanel.Dock="Top" Text="Output Devices" Style="{StaticResource SectionHeader}"/>

                            <!-- Volume + Mute row (docked to bottom) -->
                            <Grid DockPanel.Dock="Bottom" Margin="0,12,0,0">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Button x:Name="WPFSetDefaultOutput" Grid.Row="0"
                                        Style="{StaticResource BtnAccent}"
                                        Content="★  Set as Default Output" Margin="0,0,0,12"
                                        IsEnabled="False"/>
                                <Grid Grid.Row="1">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <Slider x:Name="WPFOutputVolumeSlider" Grid.Column="0"
                                            Style="{StaticResource AudioSlider}" IsEnabled="False"/>
                                    <TextBlock x:Name="WPFOutputVolumeLabel" Grid.Column="1"
                                               Text="--%" Foreground="#E0E0E0" FontWeight="SemiBold"
                                               Width="40" Margin="10,0" VerticalAlignment="Center"/>
                                    <ToggleButton x:Name="WPFOutputMuteButton" Grid.Column="2"
                                                  Style="{StaticResource MuteBtn}" Content="🔇"
                                                  IsEnabled="False" ToolTip="Mute output device"/>
                                </Grid>
                            </Grid>

                            <!-- Device list -->
                            <ListBox x:Name="WPFOutputDeviceList" Margin="0,0,0,12"/>
                        </DockPanel>
                    </Border>

                    <!-- INPUT -->
                    <Border Grid.Column="2" Style="{StaticResource Card}">
                        <DockPanel>
                            <TextBlock DockPanel.Dock="Top" Text="Input Devices" Style="{StaticResource SectionHeader}"/>

                            <Grid DockPanel.Dock="Bottom" Margin="0,12,0,0">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Button x:Name="WPFSetDefaultInput" Grid.Row="0"
                                        Style="{StaticResource BtnAccent}"
                                        Content="★  Set as Default Input" Margin="0,0,0,12"
                                        IsEnabled="False"/>
                                <Grid Grid.Row="1">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <Slider x:Name="WPFInputVolumeSlider" Grid.Column="0"
                                            Style="{StaticResource AudioSlider}" IsEnabled="False"/>
                                    <TextBlock x:Name="WPFInputVolumeLabel" Grid.Column="1"
                                               Text="--%" Foreground="#E0E0E0" FontWeight="SemiBold"
                                               Width="40" Margin="10,0" VerticalAlignment="Center"/>
                                    <ToggleButton x:Name="WPFInputMuteButton" Grid.Column="2"
                                                  Style="{StaticResource MuteBtn}" Content="🎤"
                                                  IsEnabled="False" ToolTip="Mute input device"/>
                                </Grid>
                            </Grid>

                            <ListBox x:Name="WPFInputDeviceList" Margin="0,0,0,12"/>
                        </DockPanel>
                    </Border>
                </Grid>
            </TabItem>

            <!-- ── TAB 2: APPLICATIONS ───────────────────────────────────── -->
            <TabItem x:Name="WPFTab2">
                <DockPanel Margin="20">
                    <Border DockPanel.Dock="Top" Style="{StaticResource Card}" Margin="0,0,0,16">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Per-Application Volume Control" Style="{StaticResource SectionHeader}"
                                       Margin="0" VerticalAlignment="Center"/>
                            <Button x:Name="WPFRefreshApps" Style="{StaticResource BtnBase}"
                                    Content="↺  Refresh Apps" Margin="16,0,0,0" VerticalAlignment="Center"/>
                            <TextBlock x:Name="WPFAppCount" Text="" Foreground="#8892B0" FontSize="11"
                                       Margin="12,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>

                    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                        <ItemsControl x:Name="WPFAppSessionList" Margin="0">
                            <ItemsControl.ItemTemplate>
                                <DataTemplate>
                                    <Border Style="{StaticResource Card}" Margin="0,0,0,8">
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="36"/>
                                                <ColumnDefinition Width="180"/>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="50"/>
                                                <ColumnDefinition Width="70"/>
                                            </Grid.ColumnDefinitions>

                                            <TextBlock Grid.Column="0" Text="{Binding Icon}" FontSize="20"
                                                       VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="8,0">
                                                <TextBlock Text="{Binding Name}" FontWeight="SemiBold"
                                                           Foreground="#E0E0E0" TextTrimming="CharacterEllipsis"/>
                                                <TextBlock Text="{Binding PidLabel}" Foreground="#8892B0" FontSize="10"/>
                                            </StackPanel>
                                            <Slider Grid.Column="2" Value="{Binding VolumePercent, Mode=TwoWay}"
                                                    Style="{StaticResource AudioSlider}" VerticalAlignment="Center"
                                                    Tag="{Binding SessionKey}"/>
                                            <TextBlock Grid.Column="3" Text="{Binding VolumeLabel}"
                                                       Foreground="#E0E0E0" FontWeight="SemiBold"
                                                       VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                            <ToggleButton Grid.Column="4" IsChecked="{Binding IsMuted, Mode=TwoWay}"
                                                          Style="{StaticResource MuteBtn}" Content="🔇"
                                                          Tag="{Binding SessionKey}" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                </DataTemplate>
                            </ItemsControl.ItemTemplate>
                        </ItemsControl>
                    </ScrollViewer>
                </DockPanel>
            </TabItem>

            <!-- ── TAB 3: FORMATS ────────────────────────────────────────── -->
            <TabItem x:Name="WPFTab3">
                <Grid Margin="20">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- OUTPUT FORMAT -->
                    <Border Grid.Column="0" Style="{StaticResource Card}">
                        <StackPanel>
                            <TextBlock Text="Output Format" Style="{StaticResource SectionHeader}"/>

                            <TextBlock Text="Device" Style="{StaticResource Label}"/>
                            <ComboBox x:Name="WPFFormatOutputDevice" Margin="0,0,0,12"/>

                            <TextBlock Text="Current Format" Style="{StaticResource Label}"/>
                            <Border Background="#0F1729" BorderBrush="#2D3561" BorderThickness="1"
                                    CornerRadius="4" Padding="8,6" Margin="0,0,0,16">
                                <TextBlock x:Name="WPFCurrentOutputFormat" Text="—"
                                           Foreground="#43D97B" FontFamily="Consolas" FontSize="12"/>
                            </Border>

                            <TextBlock Text="Sample Rate (Hz)" Style="{StaticResource Label}"/>
                            <ComboBox x:Name="WPFOutputSampleRate" Margin="0,0,0,12">
                                <ComboBoxItem Content="44100" Tag="44100"/>
                                <ComboBoxItem Content="48000" Tag="48000" IsSelected="True"/>
                                <ComboBoxItem Content="88200" Tag="88200"/>
                                <ComboBoxItem Content="96000" Tag="96000"/>
                                <ComboBoxItem Content="176400" Tag="176400"/>
                                <ComboBoxItem Content="192000" Tag="192000"/>
                            </ComboBox>

                            <TextBlock Text="Bit Depth" Style="{StaticResource Label}"/>
                            <ComboBox x:Name="WPFOutputBitDepth" Margin="0,0,0,16">
                                <ComboBoxItem Content="16-bit" Tag="16"/>
                                <ComboBoxItem Content="24-bit" Tag="24" IsSelected="True"/>
                                <ComboBoxItem Content="32-bit (float)" Tag="32"/>
                            </ComboBox>

                            <Button x:Name="WPFApplyOutputFormat" Style="{StaticResource BtnAccent}"
                                    Content="Apply Output Format" Margin="0,0,0,20" IsEnabled="False"/>

                            <Separator Background="#2D3561" Margin="0,0,0,16"/>

                            <TextBlock Text="Audio Enhancements" Style="{StaticResource SectionHeader}"/>
                            <TextBlock Text="Disable system-level audio enhancements for this device."
                                       Style="{StaticResource Label}" TextWrapping="Wrap" Margin="0,0,0,10"/>
                            <ToggleButton x:Name="WPFOutputEnhancementsToggle" Style="{StaticResource MuteBtn}"
                                          Content="Enhancements: Enabled" MinWidth="180" IsEnabled="False"
                                          HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Border>

                    <!-- INPUT FORMAT -->
                    <Border Grid.Column="2" Style="{StaticResource Card}">
                        <StackPanel>
                            <TextBlock Text="Input Format" Style="{StaticResource SectionHeader}"/>

                            <TextBlock Text="Device" Style="{StaticResource Label}"/>
                            <ComboBox x:Name="WPFFormatInputDevice" Margin="0,0,0,12"/>

                            <TextBlock Text="Current Format" Style="{StaticResource Label}"/>
                            <Border Background="#0F1729" BorderBrush="#2D3561" BorderThickness="1"
                                    CornerRadius="4" Padding="8,6" Margin="0,0,0,16">
                                <TextBlock x:Name="WPFCurrentInputFormat" Text="—"
                                           Foreground="#43D97B" FontFamily="Consolas" FontSize="12"/>
                            </Border>

                            <TextBlock Text="Sample Rate (Hz)" Style="{StaticResource Label}"/>
                            <ComboBox x:Name="WPFInputSampleRate" Margin="0,0,0,12">
                                <ComboBoxItem Content="8000"  Tag="8000"/>
                                <ComboBoxItem Content="16000" Tag="16000"/>
                                <ComboBoxItem Content="44100" Tag="44100"/>
                                <ComboBoxItem Content="48000" Tag="48000" IsSelected="True"/>
                                <ComboBoxItem Content="96000" Tag="96000"/>
                            </ComboBox>

                            <TextBlock Text="Bit Depth" Style="{StaticResource Label}"/>
                            <ComboBox x:Name="WPFInputBitDepth" Margin="0,0,0,16">
                                <ComboBoxItem Content="16-bit" Tag="16" IsSelected="True"/>
                                <ComboBoxItem Content="24-bit" Tag="24"/>
                                <ComboBoxItem Content="32-bit (float)" Tag="32"/>
                            </ComboBox>

                            <Button x:Name="WPFApplyInputFormat" Style="{StaticResource BtnAccent}"
                                    Content="Apply Input Format" Margin="0,0,0,20" IsEnabled="False"/>

                            <Separator Background="#2D3561" Margin="0,0,0,16"/>

                            <TextBlock Text="Audio Enhancements" Style="{StaticResource SectionHeader}"/>
                            <TextBlock Text="Disable system-level audio enhancements for this device."
                                       Style="{StaticResource Label}" TextWrapping="Wrap" Margin="0,0,0,10"/>
                            <ToggleButton x:Name="WPFInputEnhancementsToggle" Style="{StaticResource MuteBtn}"
                                          Content="Enhancements: Enabled" MinWidth="180" IsEnabled="False"
                                          HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>

            <!-- ── TAB 4: PROFILES ───────────────────────────────────────── -->
            <TabItem x:Name="WPFTab4">
                <Grid Margin="20">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Saved profiles -->
                    <Border Grid.Column="0" Style="{StaticResource Card}">
                        <DockPanel>
                            <TextBlock DockPanel.Dock="Top" Text="Saved Profiles" Style="{StaticResource SectionHeader}"/>
                            <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="0,12,0,0">
                                <Button x:Name="WPFRestoreProfile" Style="{StaticResource BtnAccent}"
                                        Content="▶  Restore" IsEnabled="False" Margin="0,0,8,0"/>
                                <Button x:Name="WPFDeleteProfile" Style="{StaticResource BtnBase}"
                                        Content="🗑  Delete" IsEnabled="False"/>
                            </StackPanel>
                            <TextBlock DockPanel.Dock="Bottom" x:Name="WPFProfileInfo" Text=""
                                       Foreground="#8892B0" FontSize="11" Margin="0,8,0,0"
                                       TextWrapping="Wrap"/>
                            <ListBox x:Name="WPFProfileList"/>
                        </DockPanel>
                    </Border>

                    <!-- Save profile -->
                    <Border Grid.Column="2" Style="{StaticResource Card}">
                        <StackPanel>
                            <TextBlock Text="Save Current State" Style="{StaticResource SectionHeader}"/>

                            <TextBlock Text="Profile Name" Style="{StaticResource Label}"/>
                            <TextBox x:Name="WPFProfileNameInput" Margin="0,0,0,16"
                                     xml:space="preserve"/>

                            <TextBlock Text="Capture" Style="{StaticResource Label}"/>
                            <CheckBox x:Name="WPFProfileSaveOutputDevice"  Content="Default Output Device"
                                      IsChecked="True" Foreground="#E0E0E0" Margin="0,4"/>
                            <CheckBox x:Name="WPFProfileSaveInputDevice"   Content="Default Input Device"
                                      IsChecked="True" Foreground="#E0E0E0" Margin="0,4"/>
                            <CheckBox x:Name="WPFProfileSaveOutputVolume"  Content="Output Volume &amp; Mute"
                                      IsChecked="True" Foreground="#E0E0E0" Margin="0,4"/>
                            <CheckBox x:Name="WPFProfileSaveInputVolume"   Content="Input Volume &amp; Mute"
                                      IsChecked="True" Foreground="#E0E0E0" Margin="0,4"/>
                            <CheckBox x:Name="WPFProfileSaveAppVolumes"    Content="Per-App Volumes"
                                      IsChecked="True" Foreground="#E0E0E0" Margin="0,4,0,16"/>

                            <Button x:Name="WPFSaveProfile" Style="{StaticResource BtnAccent}"
                                    Content="💾  Save Profile"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>

        </TabControl>
    </DockPanel>
</Window>

'@

# ─── Parse XAML ──────────────────────────────────────────────────────────────

$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' `
                       -replace 'xmlns:d="[^"]*"', '' `
                       -replace 'xmlns:mc="[^"]*"', '' `
                       -replace "x:Class=`"[^`"]*`"", ''

[xml]$xaml   = $inputXML
$reader      = [System.Xml.XmlNodeReader]::new($xaml)
$sync.Form   = [Windows.Markup.XamlReader]::Load($reader)

# ─── Bind all named controls into $sync ──────────────────────────────────────

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    $ctrlName = $_.Name
    $ctrl     = $sync.Form.FindName($ctrlName)
    if ($ctrl) { $sync[$ctrlName] = $ctrl }
}

# ─── Version label ───────────────────────────────────────────────────────────

$sync.WPFVersionLabel.Text = "v$($sync.Version)"

# ─── Tab navigation ───────────────────────────────────────────────────────────

foreach ($tabBtn in @('WPFTab1BT','WPFTab2BT','WPFTab3BT','WPFTab4BT')) {
    $btnName = $tabBtn   # capture for closure
    $sync[$btnName].Add_Click({
        Invoke-WPFTab -ClickedTab $btnName
    }.GetNewClosure())
}

# ─── Master volume slider ─────────────────────────────────────────────────────

$sync.WPFMasterVolumeSlider.Add_ValueChanged({
    $pct  = [math]::Round($sync.WPFMasterVolumeSlider.Value)
    $sync.WPFMasterVolumeLabel.Text = "$pct%"
    # Debounce — only call API when user releases
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
    $sync.WPFMasterMuteButton.Content = if ($muted) { "🔇" } else { "🔊" }
})

# ─── Output device list selection ────────────────────────────────────────────

$sync.WPFOutputDeviceList.Add_SelectionChanged({
    $selected = $sync.WPFOutputDeviceList.SelectedItem
    if (-not $selected) {
        $sync.WPFSetDefaultOutput.IsEnabled    = $false
        $sync.WPFOutputVolumeSlider.IsEnabled  = $false
        $sync.WPFOutputMuteButton.IsEnabled    = $false
        return
    }
    $sync.SelectedOutputId = $selected.Tag
    $sync.WPFSetDefaultOutput.IsEnabled   = $true
    $sync.WPFOutputVolumeSlider.IsEnabled = $true
    $sync.WPFOutputMuteButton.IsEnabled   = $true

    $dev = $sync.RenderDevices | Where-Object { $_.DeviceId -eq $selected.Tag } | Select-Object -First 1
    if ($dev) {
        $pct = [math]::Round($dev.VolumeScalar * 100)
        $sync.WPFOutputVolumeSlider.Value  = $pct
        $sync.WPFOutputVolumeLabel.Text    = "$pct%"
        $sync.WPFOutputMuteButton.IsChecked = $dev.IsMuted
        # Mirror to master header if first device
        $sync.WPFMasterVolumeSlider.Value  = $pct
        $sync.WPFMasterVolumeLabel.Text    = "$pct%"
        $sync.WPFMasterMuteButton.IsChecked = $dev.IsMuted
        $sync.WPFMasterMuteButton.Content   = if ($dev.IsMuted) { "🔇" } else { "🔊" }
    }
})

$sync.WPFOutputVolumeSlider.Add_PreviewMouseLeftButtonUp({
    $level = $sync.WPFOutputVolumeSlider.Value / 100.0
    $pct   = [math]::Round($sync.WPFOutputVolumeSlider.Value)
    $sync.WPFOutputVolumeLabel.Text     = "$pct%"
    $sync.WPFMasterVolumeSlider.Value   = $pct
    $sync.WPFMasterVolumeLabel.Text     = "$pct%"
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

# ─── Input device list selection ─────────────────────────────────────────────

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
        $sync.WPFInputVolumeSlider.Value    = $pct
        $sync.WPFInputVolumeLabel.Text      = "$pct%"
        $sync.WPFInputMuteButton.IsChecked  = $dev.IsMuted
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

# ─── Formats tab device picker change ────────────────────────────────────────

$sync.WPFFormatOutputDevice.Add_SelectionChanged({ Update-OutputFormatDisplay })
$sync.WPFFormatInputDevice.Add_SelectionChanged({  Update-InputFormatDisplay  })

# ─── Enhancements toggles ─────────────────────────────────────────────────────

$sync.WPFOutputEnhancementsToggle.Add_Click({
    $enabled  = $sync.WPFOutputEnhancementsToggle.IsChecked
    $selected = $sync.WPFFormatOutputDevice.SelectedItem
    if (-not $selected) { return }
    $deviceId = $selected.Tag
    Invoke-AudioManagerRunspace {
        $ok = Set-AudioEnhancement -DeviceId $deviceId -Enabled $enabled
        Invoke-WPFUIThread {
            $sync.WPFOutputEnhancementsToggle.Content = if ($enabled) { "Enhancements: Enabled" } else { "Enhancements: Disabled" }
            Set-WPFStatus (if ($ok) { "Output enhancements $(if ($enabled) {'enabled'} else {'disabled'})." } else { "Failed to change enhancements." })
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
            $sync.WPFInputEnhancementsToggle.Content = if ($enabled) { "Enhancements: Enabled" } else { "Enhancements: Disabled" }
            Set-WPFStatus (if ($ok) { "Input enhancements $(if ($enabled) {'enabled'} else {'disabled'})." } else { "Failed to change enhancements." })
        }
    }
})

# ─── Profile list selection ───────────────────────────────────────────────────

$sync.WPFProfileList.Add_SelectionChanged({
    $selected = $sync.WPFProfileList.SelectedItem
    $hasProfile = ($selected -and $selected.Tag)
    $sync.WPFRestoreProfile.IsEnabled = $hasProfile
    $sync.WPFDeleteProfile.IsEnabled  = $hasProfile
    if ($hasProfile) {
        $p = $selected.Tag
        $info = "Created: $($p.created)"
        if ($p.defaultOutputDeviceName) { $info += "`nOutput: $($p.defaultOutputDeviceName)" }
        if ($p.defaultInputDeviceName)  { $info += "`nInput: $($p.defaultInputDeviceName)" }
        $sync.WPFProfileInfo.Text = $info
    }
})

# ─── Button dispatcher wiring ─────────────────────────────────────────────────

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

# ─── Initial data load ────────────────────────────────────────────────────────

Set-WPFStatus "Loading audio devices..."
Invoke-WPFRefreshAll

# ─── Show window ─────────────────────────────────────────────────────────────

$sync.Form.ShowDialog() | Out-Null

# ─── Cleanup ─────────────────────────────────────────────────────────────────

$sync.RunspacePool.Close()
$sync.RunspacePool.Dispose()
Stop-Transcript -ErrorAction SilentlyContinue



