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
