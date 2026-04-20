function Initialize-DevicesTab {
    # Populate output device list
    $sync.WPFOutputDeviceList.Items.Clear()
    foreach ($dev in $sync.RenderDevices) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = $dev.Name
        $item.Tag     = $dev.DeviceId
        $sync.WPFOutputDeviceList.Items.Add($item) | Out-Null
    }

    # Populate input device list
    $sync.WPFInputDeviceList.Items.Clear()
    foreach ($dev in $sync.CaptureDevices) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = $dev.Name
        $item.Tag     = $dev.DeviceId
        $sync.WPFInputDeviceList.Items.Add($item) | Out-Null
    }
}

function Update-DevicesTab {
    Initialize-DevicesTab

    # If a device was selected, restore selection and update sliders
    if ($sync.SelectedOutputId) {
        $items = $sync.WPFOutputDeviceList.Items
        for ($i = 0; $i -lt $items.Count; $i++) {
            if ($items[$i].Tag -eq $sync.SelectedOutputId) {
                $sync.WPFOutputDeviceList.SelectedIndex = $i
                break
            }
        }
    }

    if ($sync.SelectedInputId) {
        $items = $sync.WPFInputDeviceList.Items
        for ($i = 0; $i -lt $items.Count; $i++) {
            if ($items[$i].Tag -eq $sync.SelectedInputId) {
                $sync.WPFInputDeviceList.SelectedIndex = $i
                break
            }
        }
    }
}
