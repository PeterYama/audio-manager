function Get-DeviceFormat {
    param([Parameter(Mandatory)][string]$DeviceId)
    try {
        return [AudioManager.AudioManagerHelper]::GetDeviceFormat($DeviceId)
    } catch {
        Write-Warning "Get-DeviceFormat error: $_"
        return $null
    }
}
