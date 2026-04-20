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
