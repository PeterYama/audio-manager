function Initialize-ApplicationsTab {
    $sync.WPFAppSessionList.Items.Clear()

    if ($sync.AudioSessions.Count -eq 0) {
        $sync.WPFAppCount.Text = "No active audio sessions found."
        return
    }

    $sync.WPFAppCount.Text = "$($sync.AudioSessions.Count) session(s)"

    foreach ($session in ($sync.AudioSessions | Sort-Object Name)) {
        $sync.WPFAppSessionList.Items.Add($session) | Out-Null
    }
}

function Update-ApplicationsTab {
    Initialize-ApplicationsTab
}
