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
