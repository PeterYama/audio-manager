function Initialize-FormatsTab {
    # Populate output device picker
    $sync.WPFFormatOutputDevice.Items.Clear()
    foreach ($dev in $sync.RenderDevices) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $dev.Name
        $item.Tag     = $dev.DeviceId
        $sync.WPFFormatOutputDevice.Items.Add($item) | Out-Null
    }
    if ($sync.WPFFormatOutputDevice.Items.Count -gt 0) {
        $sync.WPFFormatOutputDevice.SelectedIndex = 0
        Update-OutputFormatDisplay
    }

    # Populate input device picker
    $sync.WPFFormatInputDevice.Items.Clear()
    foreach ($dev in $sync.CaptureDevices) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $dev.Name
        $item.Tag     = $dev.DeviceId
        $sync.WPFFormatInputDevice.Items.Add($item) | Out-Null
    }
    if ($sync.WPFFormatInputDevice.Items.Count -gt 0) {
        $sync.WPFFormatInputDevice.SelectedIndex = 0
        Update-InputFormatDisplay
    }
}

function Update-OutputFormatDisplay {
    $selected = $sync.WPFFormatOutputDevice.SelectedItem
    if (-not $selected) { return }
    $deviceId = $selected.Tag
    $fmt = Get-DeviceFormat -DeviceId $deviceId
    if ($fmt) {
        $sync.WPFCurrentOutputFormat.Text = $fmt.Description
        $sync.WPFApplyOutputFormat.IsEnabled = $true

        # Reflect enhancements state
        $enhanced = Get-AudioEnhancement -DeviceId $deviceId
        $sync.WPFOutputEnhancementsToggle.IsChecked = $enhanced
        $sync.WPFOutputEnhancementsToggle.Content   = if ($enhanced) { "Enhancements: Enabled" } else { "Enhancements: Disabled" }
        $sync.WPFOutputEnhancementsToggle.IsEnabled  = $true
    } else {
        $sync.WPFCurrentOutputFormat.Text = "Unable to read format"
    }
}

function Update-InputFormatDisplay {
    $selected = $sync.WPFFormatInputDevice.SelectedItem
    if (-not $selected) { return }
    $deviceId = $selected.Tag
    $fmt = Get-DeviceFormat -DeviceId $deviceId
    if ($fmt) {
        $sync.WPFCurrentInputFormat.Text = $fmt.Description
        $sync.WPFApplyInputFormat.IsEnabled = $true

        $enhanced = Get-AudioEnhancement -DeviceId $deviceId
        $sync.WPFInputEnhancementsToggle.IsChecked = $enhanced
        $sync.WPFInputEnhancementsToggle.Content   = if ($enhanced) { "Enhancements: Enabled" } else { "Enhancements: Disabled" }
        $sync.WPFInputEnhancementsToggle.IsEnabled  = $true
    } else {
        $sync.WPFCurrentInputFormat.Text = "Unable to read format"
    }
}
