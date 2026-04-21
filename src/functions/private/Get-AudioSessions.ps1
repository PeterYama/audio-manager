function Get-AudioSessions {
    try {
        return [AudioManager.AudioManagerHelper]::GetAudioSessions()
    } catch {
        Write-Warning "Get-AudioSessions error: $_"
        return @()
    }
}
