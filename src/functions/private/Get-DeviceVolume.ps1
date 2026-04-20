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
