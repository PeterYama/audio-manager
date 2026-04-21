function Set-DeviceMute {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][bool]$Muted
    )
    return [AudioManager.AudioManagerHelper]::SetDeviceMute($DeviceId, $Muted)
}
