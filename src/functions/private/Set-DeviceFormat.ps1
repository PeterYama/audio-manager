function Set-DeviceFormat {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][int]$SampleRate,
        [Parameter(Mandatory)][int]$BitDepth,
        [int]$Channels = 2
    )
    return [AudioManager.AudioManagerHelper]::SetDeviceFormat($DeviceId, $SampleRate, $BitDepth, $Channels)
}
