function Get-AudioDevices {
    $result = @{ Render = @(); Capture = @() }
    try {
        $result.Render  = [AudioManager.AudioManagerHelper]::GetRenderDevices()
        $result.Capture = [AudioManager.AudioManagerHelper]::GetCaptureDevices()
    } catch {
        Write-Warning "Get-AudioDevices error: $_"
    }
    return $result
}
