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
