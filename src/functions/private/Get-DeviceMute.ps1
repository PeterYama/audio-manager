function Get-DeviceMute {
    param([Parameter(Mandatory)][string]$DeviceId)
    $devices = [AudioManager.AudioManagerHelper]::GetRenderDevices() +
               [AudioManager.AudioManagerHelper]::GetCaptureDevices()
    $dev = $devices | Where-Object { $_.DeviceId -eq $DeviceId } | Select-Object -First 1
    return if ($dev) { $dev.IsMuted } else { $false }
}
