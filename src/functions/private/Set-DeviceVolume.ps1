function Set-DeviceVolume {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][float]$Level     # 0.0 - 1.0
    )
    $Level = [math]::Max(0.0, [math]::Min(1.0, $Level))
    return [AudioManager.AudioManagerHelper]::SetDeviceVolume($DeviceId, $Level)
}
