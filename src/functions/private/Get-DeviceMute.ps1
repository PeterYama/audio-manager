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
