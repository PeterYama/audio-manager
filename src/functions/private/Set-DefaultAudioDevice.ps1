function Set-DefaultAudioDevice {
    param([Parameter(Mandatory)][string]$DeviceId)
    return [AudioManager.AudioManagerHelper]::SetDefaultDevice($DeviceId)
}
