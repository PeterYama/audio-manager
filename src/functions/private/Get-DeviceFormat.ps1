function Get-DeviceFormat {
    param([Parameter(Mandatory)][string]$DeviceId)
    try {
        $policy    = [AudioManager.AudioManagerHelper]::CreatePolicyConfig()
        $fmtPtr    = [IntPtr]::Zero
        $policy.GetDeviceFormat($DeviceId, $false, [ref]$fmtPtr) | Out-Null

        if ($fmtPtr -eq [IntPtr]::Zero) { return $null }

        $fmt = [System.Runtime.InteropServices.Marshal]::PtrToStructure(
            $fmtPtr,
            [AudioManager.WAVEFORMATEX]
        )

        return [PSCustomObject]@{
            SampleRate  = $fmt.nSamplesPerSec
            BitDepth    = $fmt.wBitsPerSample
            Channels    = $fmt.nChannels
            FormatTag   = $fmt.wFormatTag
            Description = "$($fmt.nSamplesPerSec) Hz / $($fmt.wBitsPerSample)-bit / $($fmt.nChannels)ch"
        }
    } catch {
        Write-Warning "Get-DeviceFormat error: $_"
        return $null
    }
}
