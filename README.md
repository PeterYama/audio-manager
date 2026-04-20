# 🔊 Audio Manager

A PowerShell + WPF Windows app that consolidates every audio setting in one place — no installation required.

## One-liner launch

```powershell
iwr -useb https://raw.githubusercontent.com/PeterYama/audio-manager/main/AudioManager.ps1 | iex
```

> **Requires administrator privileges.** The script will automatically prompt for elevation via UAC.

---

## Features

| Tab | What you can do |
|-----|----------------|
| **Devices** | List all output/input devices, set default, adjust volume, mute/unmute per device |
| **Applications** | Per-app volume control and mute for every running audio session (like the Windows Volume Mixer, but better) |
| **Formats** | View and change sample rate (44.1–192 kHz) and bit depth (16/24/32-bit) per device; toggle audio enhancements |
| **Profiles** | Save the full state (default devices, volumes, per-app volumes) as a named profile and restore with one click |

**Header:** Master volume slider and mute always visible regardless of active tab.

---

## Requirements

- Windows 10 or 11
- PowerShell 5.1 (built-in, no install needed)
- Administrator rights (needed to change default devices and audio formats)

---

## Run locally (from source)

```powershell
git clone https://github.com/PeterYama/audio-manager
cd audio-manager
.\Compile.ps1          # Regenerates AudioManager.ps1 from src/
.\AudioManager.ps1     # Launch the app
```

---

## Project structure

```
audio-manager/
├── Compile.ps1                  # Build script
├── AudioManager.ps1             # Compiled single-file output (run this)
└── src/
    ├── core/
    │   ├── CoreAudio.cs         # C# COM P/Invoke (Windows Core Audio API)
    │   └── Initialize-CoreAudio.ps1
    ├── functions/
    │   ├── private/             # Audio API wrappers
    │   └── public/              # WPF UI event handlers
    ├── ui/
    │   └── MainWindow.xaml      # WPF layout (all 4 tabs)
    └── scripts/
        ├── start.ps1            # Init, elevation check, runspace pool
        └── main.ps1             # WPF bootstrap, event wiring, ShowDialog
```

---

## How it works

Inspired by [ChrisTitusTech/winutil](https://github.com/ChrisTitusTech/winutil):

- `Compile.ps1` concatenates all `src/` files into the single `AudioManager.ps1`
- The compiled script is hosted on GitHub and fetched directly into memory via `iwr | iex`
- The GUI is built with WPF (PresentationFramework) using an embedded XAML definition
- All audio operations use the Windows **Core Audio API** via C# COM interop (`Add-Type`)
- Background tasks run in a runspace pool to keep the UI responsive

### Audio APIs used

| Feature | API |
|---------|-----|
| Device enumeration | `IMMDeviceEnumerator` |
| Default device switching | `IPolicyConfig.SetDefaultEndpoint` |
| Device volume/mute | `IAudioEndpointVolume` |
| Per-app volume | `IAudioSessionManager2` + `ISimpleAudioVolume` |
| Sample rate / bit depth | `IPolicyConfig.GetDeviceFormat` / `SetDeviceFormat` |
| Audio enhancements | Registry `PKEY_AudioEndpoint_Disable_SysFx` |

---

## Contributing

1. Fork the repo
2. Edit files under `src/`
3. Run `.\Compile.ps1` to regenerate `AudioManager.ps1`
4. Submit a PR — CI auto-recompiles on merge to `main`

---

## License

MIT
