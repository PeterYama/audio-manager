function Set-DeviceFormat {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][uint32]$SampleRate,
        [Parameter(Mandatory)][uint16]$BitDepth
    )
    try {
        $fmt = New-Object AudioManager.WAVEFORMATEX
        # For PCM (16/24-bit) use format tag 1; for 32-bit float use 3
        $fmt.wFormatTag      = if ($BitDepth -eq 32) { [uint16]3 } else { [uint16]1 }
        $fmt.nChannels       = [uint16]2
        $fmt.nSamplesPerSec  = $SampleRate
        $fmt.wBitsPerSample  = $BitDepth
        $fmt.nBlockAlign     = [uint16](($fmt.nChannels * $fmt.wBitsPerSample) / 8)
        $fmt.nAvgBytesPerSec = $fmt.nSamplesPerSec * $fmt.nBlockAlign
        $fmt.cbSize          = [uint16]0

        $policy = [AudioManager.AudioManagerHelper]::CreatePolicyConfig()
        $policy.SetDeviceFormat($DeviceId, [ref]$fmt, [IntPtr]::Zero) | Out-Null
        return $true
    } catch {
        Write-Warning "Set-DeviceFormat error: $_"
        return $false
    }
}
