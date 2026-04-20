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
