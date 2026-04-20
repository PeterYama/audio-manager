function Set-AppMute {
    param(
        [Parameter(Mandatory)][string]$SessionKey,
        [Parameter(Mandatory)][bool]$Muted
    )
    try {
        $session = $sync.AudioSessions | Where-Object { $_.SessionKey -eq $SessionKey } | Select-Object -First 1
        if ($session -and $session.SimpleVolume) {
            $guid = [guid]::Empty
            $session.SimpleVolume.SetMute($Muted, [ref]$guid) | Out-Null
            $session.IsMuted = $Muted
        }
    } catch {
        Write-Warning "Set-AppMute error: $_"
    }
}
