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
