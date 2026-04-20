function Initialize-ProfilesTab {
    $sync.WPFProfileList.Items.Clear()
    $sync.WPFRestoreProfile.IsEnabled = $false
    $sync.WPFDeleteProfile.IsEnabled  = $false
    $sync.WPFProfileInfo.Text         = ""

    if ($sync.Profiles.Count -eq 0) {
        $sync.WPFProfileList.Items.Add("No saved profiles") | Out-Null
        return
    }

    foreach ($p in $sync.Profiles) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = $p.name
        $item.Tag     = $p
        $item.ToolTip = "Saved: $($p.created)"
        $sync.WPFProfileList.Items.Add($item) | Out-Null
    }
}

function Update-ProfilesTab {
    Initialize-ProfilesTab
}
