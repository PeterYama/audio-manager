function Set-AppMute {
    param(
        [Parameter(Mandatory)][int]$ProcessId,
        [Parameter(Mandatory)][bool]$Muted
    )
    return [AudioManager.AudioManagerHelper]::SetSessionMute($ProcessId, $Muted)
}
