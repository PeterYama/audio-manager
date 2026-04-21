function Invoke-WPFRefreshAll {
    if ($sync.IsRefreshing) { return }
    $sync.IsRefreshing = $true
    Set-WPFStatus "Refreshing audio devices..."

    Invoke-AudioManagerRunspace {
        try {
            $devices  = Get-AudioDevices
            $sessions = Get-AudioSessions
            $profiles = Get-AudioProfiles

            $sync.RenderDevices  = $devices.Render
            $sync.CaptureDevices = $devices.Capture
            $sync.AudioSessions  = $sessions
            $sync.Profiles       = $profiles

            Invoke-WPFUIThread {
                Initialize-DevicesTab
                Initialize-ApplicationsTab
                Initialize-FormatsTab
                Initialize-ProfilesTab
                Set-WPFStatus "Ready - $($sync.RenderDevices.Count) output, $($sync.CaptureDevices.Count) input device(s) found."
            }
        } catch {
            Invoke-WPFUIThread { Set-WPFStatus "Refresh error: $_" }
        } finally {
            $sync.IsRefreshing = $false
        }
    }
}
