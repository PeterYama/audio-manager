function Set-AppVolume {
    param(
        [Parameter(Mandatory)][string]$SessionKey,   # "pid-index"
        [Parameter(Mandatory)][float]$Level           # 0.0 - 1.0
    )
    try {
        $Level = [math]::Max(0.0, [math]::Min(1.0, $Level))
        $session = $sync.AudioSessions | Where-Object { $_.SessionKey -eq $SessionKey } | Select-Object -First 1
        if ($session -and $session.SimpleVolume) {
            $guid = [guid]::Empty
            $session.SimpleVolume.SetMasterVolume($Level, [ref]$guid) | Out-Null
            $session.VolumePercent = [math]::Round($Level * 100)
            $session.VolumeLabel   = "$($session.VolumePercent)%"
        }
    } catch {
        Write-Warning "Set-AppVolume error: $_"
    }
}
